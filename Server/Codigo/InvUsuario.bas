Attribute VB_Name = "InvUsuario"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
' 22/05/2010: Los items newbies ya no son robables.
'***************************************************

'17/09/02
'Agregue que la función se asegure que el objeto no es un barco

On Error GoTo ErrHandler

    Dim I As Integer
    Dim ObjIndex As Integer
    
    For I = 1 To UserList(UserIndex).CurrentInventorySlots
        ObjIndex = UserList(UserIndex).Invent.Object(I).ObjIndex
        If ObjIndex > 0 Then
            If (ObjData(ObjIndex).ObjType <> eOBJType.otLlaves And _
                ObjData(ObjIndex).ObjType <> eOBJType.otBarcos And _
                Not ItemNewbie(ObjIndex)) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
        End If
    Next I
    
    Exit Function

ErrHandler:
    Call LogError("Error en TieneObjetosRobables. Error: " & Err.Number & " - " & Err.Description)
End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 14/01/2010 (ZaMa)
'14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
'***************************************************

On Error GoTo manejador
    
    ClasePuedeUsarItem = True
    
    'Admins can use ANYTHING!
    If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
    
        If MasteryAllowToEquipItem(UserIndex, ObjIndex) Then
            ClasePuedeUsarItem = True
            Exit Function
        End If
        
        If ObjData(ObjIndex).ClaseProhibida(1) <> 0 Then
            Dim I As Integer
            For I = 1 To NUMCLASES
                If ObjData(ObjIndex).ClaseProhibida(I) = UserList(UserIndex).clase Then
                    ClasePuedeUsarItem = False
                    Exit Function
                End If
            Next I
        End If
        
    End If

Exit Function

manejador:
    LogError ("Error para objeto " & ObjIndex & " en ClasePuedeUsarItem de InvUsuario.bas")
End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

Dim J As Integer

With UserList(UserIndex)
    For J = 1 To UserList(UserIndex).CurrentInventorySlots
        If .Invent.Object(J).ObjIndex > 0 Then
            If ObjData(.Invent.Object(J).ObjIndex).Newbie = 1 Then _
                Call QuitarUserInvItem(UserIndex, J, MAX_INVENTORY_OBJS)
        End If
    Next J
    
    '[Barrin 17-12-03] Si el usuario dejó de ser Newbie, y estaba en el Newbie Dungeon
    'es transportado a su hogar de origen ;)
    If MapInfo(.Pos.Map).Restringir = eRestrict.restrict_newbie Then
        
        Dim DeDonde As WorldPos
        
        Select Case .Hogar
            Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
                DeDonde = Lindos
            Case eCiudad.cUllathorpe
                DeDonde = Ullathorpe
            Case eCiudad.cBanderbill
                DeDonde = Banderbill
            Case Else
                DeDonde = Nix
        End Select
        
        Call WarpUserChar(UserIndex, DeDonde.Map, DeDonde.X, DeDonde.Y, True)
    
    End If
    '[/Barrin]
End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub QuitarNewbieObj de InvUsuario.bas")
End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

Dim J As Integer

With UserList(UserIndex)
    For J = 1 To .CurrentInventorySlots
        .Invent.Object(J).ObjIndex = 0
        .Invent.Object(J).Amount = 0
        .Invent.Object(J).Equipped = 0
    Next J
    
    .Invent.NroItems = 0
    
    .Invent.ArmourEqpObjIndex = 0
    .Invent.ArmourEqpSlot = 0
    
    .Invent.WeaponEqpObjIndex = 0
    .Invent.WeaponEqpSlot = 0
    
    .Invent.CascoEqpObjIndex = 0
    .Invent.CascoEqpSlot = 0
    
    .Invent.EscudoEqpObjIndex = 0
    .Invent.EscudoEqpSlot = 0
    
    .Invent.AnilloEqpObjIndex = 0
    .Invent.AnilloEqpSlot = 0
    
    .Invent.MunicionEqpObjIndex = 0
    .Invent.MunicionEqpSlot = 0
    
    .Invent.BarcoObjIndex = 0
    .Invent.BarcoSlot = 0
    
    .Invent.MochilaEqpObjIndex = 0
    .Invent.MochilaEqpSlot = 0
    
    .Invent.FactionArmourEqpObjIndex = 0
    .Invent.FactionArmourEqpSlot = 0
    
End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LimpiarInventario de InvUsuario.bas")
End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 23/01/2007
'23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
'***************************************************
On Error GoTo ErrHandler

'If Cantidad > 100000 Then Exit Sub

With UserList(UserIndex)
    'SI EL Pjta TIENE ORO LO TIRAMOS
    If (Cantidad > 0) And (Cantidad <= .Stats.GLD) Then
            Dim MiObj As Obj
            'info debug
            Dim Loops As Integer
            
            If .flags.DueloIndex > 0 Then
                Call WriteConsoleMsg(UserIndex, "No puedes tirar oro mientras tienes una petición de duelo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Seguridad Alkon (guardo el oro tirado si supera los 50k)
            If Cantidad >= MIN_GOLD_AMOUNT_LOG Then
                Dim J As Integer
                Dim k As Integer
                Dim m As Integer
                Dim Cercanos As String
                m = .Pos.Map
                For J = .Pos.X - 10 To .Pos.X + 10
                    For k = .Pos.Y - 10 To .Pos.Y + 10
                        If InMapBounds(m, J, k) Then
                            If MapData(m, J, k).UserIndex > 0 Then
                                Cercanos = Cercanos & UserList(MapData(m, J, k).UserIndex).Name & ","
                            End If
                        End If
                    Next k
                Next J
                Cercanos = Left$(Cercanos, Len(Cercanos) - 1)
                Call LogDesarrollo(.Name & " tir? " & Cantidad & " monedas de oro en " & .Pos.Map & ", " & .Pos.X & ", " & .Pos.Y & ". Cercanos: " & Cercanos)
            End If
            '/Seguridad
            Dim Extra As Long
            Dim TeniaOro As Long
            TeniaOro = .Stats.GLD
            If Cantidad > 500000 Then 'Para evitar explotar demasiado
                Extra = Cantidad - 500000
                Cantidad = 500000
            End If
            
            Do While (Cantidad > 0)
                
                If Cantidad > MAX_INVENTORY_OBJS And .Stats.GLD > MAX_INVENTORY_OBJS Then
                    MiObj.Amount = MAX_INVENTORY_OBJS
                    Cantidad = Cantidad - MiObj.Amount
                Else
                    MiObj.Amount = Cantidad
                    Cantidad = Cantidad - MiObj.Amount
                End If
    
                MiObj.ObjIndex = ConstantesItems.Oro
                
                If EsGm(UserIndex) Then Call LogGM(.Name, "Tiró cantidad:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name)
                Dim AuxPos As WorldPos
                
                If .clase = eClass.Thief And .Invent.BarcoObjIndex = 476 Then
                    AuxPos = TirarItemAlPiso(.Pos, MiObj, False)
                    If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                        .Stats.GLD = .Stats.GLD - MiObj.Amount
                    End If
                Else
                    AuxPos = TirarItemAlPiso(.Pos, MiObj, True)
                    If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                        .Stats.GLD = .Stats.GLD - MiObj.Amount
                    End If
                End If
                
                'info debug
                Loops = Loops + 1
                If Loops > 100 Then
                    LogError ("Error en tiraroro")
                    Exit Sub
                End If
                
            Loop
            If TeniaOro = .Stats.GLD Then Extra = 0
            If Extra > 0 Then
                .Stats.GLD = .Stats.GLD - Extra
            End If
        
    End If
End With

Exit Sub

ErrHandler:
    Call LogError("Error en TirarOro. Error " & Err.Number & " : " & Err.Description)
End Sub
Public Function GetUserInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte) As UserOBJ

On Error GoTo ErrHandler

    If Slot < 1 Or Slot > UserList(UserIndex).CurrentInventorySlots Then Exit Function
    
    GetUserInvItem = UserList(UserIndex).Invent.Object(Slot)

Exit Function

ErrHandler:
    Call LogError("Error en GetUserInvItem. Error " & Err.Number & " : " & Err.Description)
    
End Function
Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    If Slot < 1 Or Slot > UserList(UserIndex).CurrentInventorySlots Then Exit Sub
    
    With UserList(UserIndex).Invent.Object(Slot)
        If (.Amount <= Cantidad) And (.Equipped = 1) Then
            Call Desequipar(UserIndex, Slot, True)
        End If
        
        'Quita un objeto
        .Amount = .Amount - Cantidad
        '¿Quedan mas?
        If .Amount <= 0 Then
            UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
            .ObjIndex = 0
            .Amount = 0
        End If
        
        Call UpdateUserInv(False, UserIndex, Slot)
    End With

Exit Sub

ErrHandler:
    Call LogError("Error en QuitarUserInvItem. Error " & Err.Number & " : " & Err.Description)
    
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim ObjIndex As Integer
    Dim Amount As Integer
    Dim Equipped As Byte
    Dim CanUse As Boolean
    
    Dim LoopC As Long
    
    With UserList(UserIndex)
        'Actualiza un solo slot
        If Not UpdateAll Then
        
            'Actualiza el slot
            ObjIndex = .Invent.Object(Slot).ObjIndex
            
            If ObjIndex > 0 Then
                Amount = .Invent.Object(Slot).Amount
                Equipped = .Invent.Object(Slot).Equipped
                CanUse = General.checkCanUseItem(UserIndex, ObjIndex)
            End If
            
            Call WriteChangeInventorySlot(UserIndex, Slot, ObjIndex, Amount, Equipped, CanUse)
        Else
            ' Limpia todo
            Call WriteChangeInventorySlot(UserIndex, 0, 0, 0, 0, True)
            
            'Actualiza todos los slots
            For LoopC = 1 To .CurrentInventorySlots
            
                'Actualiza el inventario
                ObjIndex = .Invent.Object(LoopC).ObjIndex
                
                If ObjIndex > 0 Then
                    Amount = .Invent.Object(LoopC).Amount
                    Equipped = .Invent.Object(LoopC).Equipped
                    CanUse = General.checkCanUseItem(UserIndex, ObjIndex)
                    Call WriteChangeInventorySlot(UserIndex, CByte(LoopC), ObjIndex, Amount, Equipped, CanUse)
                End If
            Next LoopC
        End If
        
        Exit Sub
    End With

ErrHandler:
    Call LogError("Error en UpdateUserInv. Error " & Err.Number & " : " & Err.Description)
End Sub
Sub DropObjCloseUser(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Num As Integer, _
    ByVal X As Integer, ByVal Y As Integer, Optional ByVal isDrop As Boolean = False)
On Error GoTo ErrHandler
  
    Dim Diff As Integer
    
    Diff = Abs(UserList(UserIndex).Pos.X - X)
    
    If Diff > 1 Then
        Call DropObj(UserIndex, Slot, Num, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, isDrop)
        Exit Sub
    End If
    
    Diff = Abs(UserList(UserIndex).Pos.Y - Y)
    
    If Diff > 1 Then
        Call DropObj(UserIndex, Slot, Num, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, isDrop)
        Exit Sub
    End If
    
        Call DropObj(UserIndex, Slot, Num, UserList(UserIndex).Pos.Map, X, Y, isDrop)
    
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DropObjCloseUser de InvUsuario.bas")
End Sub
Sub DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Num As Integer, _
    ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal isDrop As Boolean = False)
'***************************************************
'Author: Unknown
'Last Modification: 11/5/2010
'11/5/2010 - ZaMa: Arreglo bug que permitia apilar mas de 10k de items.
'***************************************************
On Error GoTo ErrHandler
  

    Dim DropObj As Obj
    Dim MapObj As Obj
    Dim str As String
    
    With UserList(UserIndex)
        If Num > 0 Then
            
            DropObj.ObjIndex = .Invent.Object(Slot).ObjIndex
            
            If (ItemNewbie(DropObj.ObjIndex) And (.flags.Privilegios And PlayerType.User)) Then
                Call WriteConsoleMsg(UserIndex, "No puedes tirar objetos newbie.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            ' Users can't drop non-transferible items
             If ObjData(DropObj.ObjIndex).Intransferible = 1 Or ObjData(DropObj.ObjIndex).NoSeTira = 1 Then
                 If ((.flags.Privilegios And PlayerType.User) <> 0) Then
                     Call WriteConsoleMsg(UserIndex, "No puedes tirar este objeto.", FontTypeNames.FONTTYPE_FIGHT)
                     Exit Sub
                 End If
             End If
             
             If ObjData(DropObj.ObjIndex).ObjType = otQuest Then
                Call WriteConsoleMsg(UserIndex, "No puedes tirar este objeto.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
             End If

            If UserList(UserIndex).flags.DueloIndex > 0 Then
                If DuelData.Duelo(UserList(UserIndex).flags.DueloIndex).Drop And (.Pos.Map = 1 Or .Pos.Map = 171) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes tirar objetos mientras tienes una petición de duelo por drop.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
        
            DropObj.Amount = MinimoInt(Num, .Invent.Object(Slot).Amount)
    
            'Check objeto en el suelo
            MapObj.ObjIndex = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
            MapObj.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount
            
            If MapObj.ObjIndex = 0 Or MapObj.ObjIndex = DropObj.ObjIndex Then
            
                If MapObj.Amount = MAX_INVENTORY_OBJS Then
                    Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If DropObj.Amount + MapObj.Amount > MAX_INVENTORY_OBJS Then
                    DropObj.Amount = MAX_INVENTORY_OBJS - MapObj.Amount
                End If

                If ObjData(DropObj.ObjIndex).ObjType = eOBJType.otBarcos Then
                    If .flags.Navegando = 1 And Slot = .Invent.BarcoSlot And (.Invent.Object(Slot).Amount - DropObj.Amount) < 1 Then
                        Call WriteConsoleMsg(UserIndex, "No puedes tirar tu barca mientras navegas.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

                    Call WriteConsoleMsg(UserIndex, "¡¡ATENCIÓN!! ¡ACABAS DE TIRAR TU BARCA!", FontTypeNames.FONTTYPE_TALK)
                End If

                Call MakeObj(DropObj, Map, X, Y)
                Call QuitarUserInvItem(UserIndex, Slot, DropObj.Amount)
                
                If Not .flags.Privilegios And PlayerType.User Then Call LogGM(.Name, "Tiró cantidad:" & DropObj.Amount & " Objeto:" & ObjData(DropObj.ObjIndex).Name)
                
                'Log de Objetos que se tiran al piso. Pablo (ToxicWaste) 07/09/07
                'Es un Objeto que tenemos que loguear?
                If ObjData(DropObj.ObjIndex).Log = 1 Or (ObjData(DropObj.ObjIndex).ObjType = eOBJType.otLlaves) Then
                    Call LogDesarrollo(.Name & " tiró al piso " & IIf(isDrop, "", "al morir ") & DropObj.Amount & " " & ObjData(DropObj.ObjIndex).Name & " Mapa: " & Map & " X: " & X & " Y: " & Y)
                ElseIf DropObj.Amount > MIN_AMOUNT_LOG Then  'Es mucha cantidad? > Subí a 5000 el minimo porque si no se llenaba el log de cosas al pedo. (NicoNZ)
                    'Si no es de los prohibidos de loguear, lo logueamos.
                    If ObjData(DropObj.ObjIndex).NoLog <> 1 Then
                        Call LogDesarrollo(.Name & " tiró al piso " & IIf(isDrop, "", "al morir ") & DropObj.Amount & " " & ObjData(DropObj.ObjIndex).Name & " Mapa: " & Map & " X: " & X & " Y: " & Y)
                    End If
                    
                ElseIf (DropObj.Amount * ObjData(DropObj.ObjIndex).Valor) >= MIN_VALUE_LOG Then
                    'Si no es de los prohibidos de loguear, lo logueamos.
                    If ObjData(DropObj.ObjIndex).NoLog <> 1 Then
                        Call LogDesarrollo(.Name & " tiró al piso " & IIf(isDrop, "", "al morir ") & DropObj.Amount & " " & ObjData(DropObj.ObjIndex).Name & " Mapa: " & Map & " X: " & X & " Y: " & Y)
                    End If
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DropObj de InvUsuario.bas")
End Sub

Sub EraseObj(ByVal Num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

With MapData(Map, X, Y)
    .ObjInfo.Amount = .ObjInfo.Amount - Num
    
    If .ObjInfo.Amount <= 0 Then
        .ObjInfo.ObjIndex = 0
        .ObjInfo.Amount = 0
        .ObjInfo.ActivatedByUser = 0
        .ObjInfo.PendingQty = 0
        
        Call ModAreas.DeleteEntity(ModAreas.Pack(Map, X, Y), ENTITY_TYPE_OBJECT)
    End If
End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EraseObj de InvUsuario.bas")
End Sub

Sub MakeObj(ByRef Obj As Obj, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  
    
    If Obj.ObjIndex > 0 And Obj.ObjIndex <= UBound(ObjData) Then
        With MapData(Map, X, Y)
            .ObjInfo.CurrentGrhIndex = ObjData(Obj.ObjIndex).GrhIndex
            Obj.CurrentGrhIndex = .ObjInfo.CurrentGrhIndex
            
            If .ObjInfo.ObjIndex = Obj.ObjIndex Then
                .ObjInfo.Amount = .ObjInfo.Amount + Obj.Amount
                
            Else
                .ObjInfo = Obj
                
                Dim Coordinates As WorldPos
                Coordinates.Map = Map
                Coordinates.X = X
                Coordinates.Y = Y
                
                If .ObjInfo.ObjIndex > 0 Then
                    Call ModAreas.DeleteEntity(ModAreas.Pack(Map, X, Y), ENTITY_TYPE_OBJECT)
                End If
                
                Call ModAreas.CreateEntity(ModAreas.Pack(Map, X, Y), ENTITY_TYPE_OBJECT, Coordinates, ObjData(.ObjInfo.ObjIndex).SizeWidth, ObjData(.ObjInfo.ObjIndex).SizeHeight)
            End If
        End With
    End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub MakeObj de InvUsuario.bas")
End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As Obj) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler
    Dim Slot As Byte
    
    With UserList(UserIndex)
        '¿el user ya tiene un objeto del mismo tipo?
        Slot = 1
        
        Do Until .Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And _
                 .Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
           Slot = Slot + 1
           If Slot > .CurrentInventorySlots Then
                 Exit Do
           End If
        Loop
            
        'Sino busca un slot vacio
        If Slot > .CurrentInventorySlots Then
           Slot = 1
           Do Until .Invent.Object(Slot).ObjIndex = 0
               Slot = Slot + 1
               If Slot > .CurrentInventorySlots Then
                   Call WriteConsoleMsg(UserIndex, "No puedes cargar más objetos.", FontTypeNames.FONTTYPE_FIGHT)
                   MeterItemEnInventario = False
                   Exit Function
               End If
           Loop
           .Invent.NroItems = .Invent.NroItems + 1
        End If
    
        If Slot > MAX_NORMAL_INVENTORY_SLOTS And Slot <= MAX_INVENTORY_SLOTS Then
            If Not ItemSeCae(MiObj.ObjIndex) Then
                Call WriteConsoleMsg(UserIndex, "No puedes contener objetos especiales en tu " & ObjData(.Invent.MochilaEqpObjIndex).Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                MeterItemEnInventario = False
                Exit Function
            End If
        End If
        'Mete el objeto
        If .Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
           'Menor que MAX_INV_OBJS
           .Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
           .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount + MiObj.Amount
        Else
           .Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS
        End If
        
        Call UpdateUserInv(False, UserIndex, Slot)
    End With
    
    MeterItemEnInventario = True
           
    Exit Function
ErrHandler:
    Call LogError("Error en MeterItemEnInventario. Error " & Err.Number & " : " & Err.Description)
End Function


Function GetObjPosition(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Integer
On Error GoTo ErrHandler

    Dim Slot As Byte
    
    With UserList(UserIndex)
        For Slot = 1 To MAX_INVENTORY_SLOTS
            If .Invent.Object(Slot).ObjIndex = ObjIndex Then
                GetObjPosition = Slot
                Exit Function
            End If
        Next Slot
    End With
    
    GetObjPosition = 0
           
    Exit Function
ErrHandler:
    Call LogError("Error en GetObjPosition. Error " & Err.Number & " : " & Err.Description)
End Function

Sub GetObj(ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 25/01/2015
'18/12/2009: ZaMa - Oro directo a la billetera.
'25/01/2015: D'Artagnan - Can't pick up items if dead.
'***************************************************
On Error GoTo ErrHandler
  

    Dim Obj As ObjData
    Dim MiObj As Obj
    Dim ObjPos As String
    
    With UserList(UserIndex)
        If .flags.Muerto = 1 Then Exit Sub
        
        '¿Hay algun obj?
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex > 0 Then
            '¿Esta permitido agarrar este obj?
            If ObjData(MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex).Agarrable <> 1 Then
                Dim X As Integer
                Dim Y As Integer
                Dim Difference As Integer, QtyLiftUp As Integer
                
                X = .Pos.X
                Y = .Pos.Y
                
                Obj = ObjData(MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex)
                MiObj.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount
                MiObj.ObjIndex = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                
                ' Oro directo a la billetera!
                If Obj.ObjType = otGuita Then
                    If .Stats.GLD >= ConstantesBalance.MaxOro Then ' If the user reached the gold limit, we do not let him lift
                        Call WriteConsoleMsg(UserIndex, "Has alcanzado el limite de oro.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    If .Stats.GLD + MiObj.Amount >= ConstantesBalance.MaxOro Then 'just pick up the difference and leave the rest on the floor
                        Difference = .Stats.GLD + MiObj.Amount - ConstantesBalance.MaxOro
                        .Stats.GLD = ConstantesBalance.MaxOro
                        QtyLiftUp = MiObj.Amount - Difference
                        MapData(.Pos.Map, X, Y).ObjInfo.Amount = Difference
                    Else
                        .Stats.GLD = .Stats.GLD + MiObj.Amount
                        QtyLiftUp = MiObj.Amount
                        'Quitamos el objeto
                        Call EraseObj(MapData(.Pos.Map, X, Y).ObjInfo.Amount, .Pos.Map, .Pos.X, .Pos.Y)
                    End If
                       
                    Call WriteUpdateGold(UserIndex)
                    
                    If QtyLiftUp >= MIN_GOLD_AMOUNT_LOG Then
                        ObjPos = " Mapa: " & .Pos.Map & " X: " & .Pos.X & " Y: " & .Pos.Y
                        Call LogDesarrollo(.Name & " juntó del piso " & QtyLiftUp & " monedas de oro" & ObjData(MiObj.ObjIndex).Name & ObjPos)
                    End If
                Else
                    If MeterItemEnInventario(UserIndex, MiObj) Then
                        
                        'Quitamos el objeto
                        Call EraseObj(MapData(.Pos.Map, X, Y).ObjInfo.Amount, .Pos.Map, .Pos.X, .Pos.Y)
                        If ((.flags.Privilegios And PlayerType.User) = 0) Then Call LogGM(.Name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name)
        
                        'Log de Objetos que se agarran del piso. Pablo (ToxicWaste) 07/09/07
                        'Es un Objeto que tenemos que loguear?
                        If ObjData(MiObj.ObjIndex).Log = 1 Or (ObjData(MiObj.ObjIndex).ObjType = eOBJType.otLlaves) Then
                            ObjPos = " Mapa: " & .Pos.Map & " X: " & .Pos.X & " Y: " & .Pos.Y
                            Call LogDesarrollo(.Name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & ObjPos)
                        ElseIf MiObj.Amount > MIN_AMOUNT_LOG Then 'Es mucha cantidad?
                            'Si no es de los prohibidos de loguear, lo logueamos.
                            If ObjData(MiObj.ObjIndex).NoLog <> 1 Then
                                ObjPos = " Mapa: " & .Pos.Map & " X: " & .Pos.X & " Y: " & .Pos.Y
                                Call LogDesarrollo(.Name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & ObjPos)
                            End If
                        ElseIf (MiObj.Amount * ObjData(MiObj.ObjIndex).Valor) >= MIN_VALUE_LOG Then
                            'Si no es de los prohibidos de loguear, lo logueamos.
                            If ObjData(MiObj.ObjIndex).NoLog <> 1 Then
                                ObjPos = " Mapa: " & .Pos.Map & " X: " & .Pos.X & " Y: " & .Pos.Y
                                Call LogDesarrollo(.Name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & ObjPos)
                            End If
                        End If
                    End If
                End If
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "No hay nada aquí.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GetObj de InvUsuario.bas")
End Sub

Public Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal RefreshChar As Boolean)
'***************************************************
'Author: Unknown
'Last Modification: 21/02/2014 (D'Artagnan)
'26/05/2011: Amraphen - Agregadas armaduras faccionarias de segunda jerarquía.
'21/02/2014: D'Artagnan - Bug fixes and code optimization.
'***************************************************

On Error GoTo ErrHandler

    'Desequipa el item slot del inventario
    Dim Obj As ObjData
    
    With UserList(UserIndex)
        With .Invent
            If (Slot < LBound(.Object)) Or (Slot > UBound(.Object)) Then
                Exit Sub
            ElseIf .Object(Slot).ObjIndex = 0 Then
                Exit Sub
            End If
            
            Obj = ObjData(.Object(Slot).ObjIndex)
        End With
        
        Select Case Obj.ObjType
            Case eOBJType.otWeapon, eOBJType.otTool
                With .Invent
                    .Object(Slot).Equipped = 0
                    .WeaponEqpObjIndex = 0
                    .WeaponEqpSlot = 0
                End With
                
                If .flags.Mimetizado = 0 Then
                    With .Char
                        .WeaponAnim = ConstantesGRH.NingunArma
                        
                        If RefreshChar And UserList(UserIndex).flags.Navegando <> 1 Then
                            Call ChangeUserChar(UserIndex, .body, .head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
                        End If
                    End With
                End If
            
            Case eOBJType.otFlechas
                With .Invent
                    .Object(Slot).Equipped = 0
                    .MunicionEqpObjIndex = 0
                    .MunicionEqpSlot = 0
                End With
            
            Case eOBJType.otAnillo
                With .Invent
                    .Object(Slot).Equipped = 0
                    .AnilloEqpObjIndex = 0
                    .AnilloEqpSlot = 0
                End With
            
            Case eOBJType.otArmadura
                'Nos fijamos si es armadura de segunda jerarquía
                If (Obj.Real = 2) Or (Obj.Caos = 2) Then
                    With .Invent
                        .Object(Slot).Equipped = 0
                        .FactionArmourEqpObjIndex = 0
                        .FactionArmourEqpSlot = 0
                    End With
                    
                    'Cambiamos el body si tiene una armadura faccionaria de defensa alta, sino no pasa nada
                    If .flags.Navegando = 0 And _
                       (.Invent.ArmourEqpObjIndex = ArmadurasFaccion(.clase, .raza).Armada(eTipoDefArmors.ieMax) Or _
                        .Invent.ArmourEqpObjIndex = ArmadurasFaccion(.clase, .raza).Armada(eTipoDefArmors.ieAlta) Or _
                        .Invent.ArmourEqpObjIndex = ArmadurasFaccion(.clase, .raza).Caos(eTipoDefArmors.ieMax) Or _
                        .Invent.ArmourEqpObjIndex = ArmadurasFaccion(.clase, .raza).Caos(eTipoDefArmors.ieAlta)) Then
                        
                        If .flags.Mimetizado Then
                            .OrigChar.body = GetBodyForUser(UserIndex, .Invent.ArmourEqpObjIndex)
                        Else
                            .Char.body = GetBodyForUser(UserIndex, .Invent.ArmourEqpObjIndex)
                            
                            With .Char
                                Call ChangeUserChar(UserIndex, .body, .head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
                            End With
                        End If
                    End If
                    
                    'TODO: Antes también actualizaba el slot de la armadura común, la verdad no sé por qué ya que no se modifica la cantidad ni si está equipada/desequipada..
                Else
                    With .Invent
                        'Si tiene armadura faccionaria de segunda jerarquía equipada la sacamos:
                        If .FactionArmourEqpObjIndex > 0 Then
                            .Object(.FactionArmourEqpSlot).Equipped = 0
                            
                            Call UpdateUserInv(False, UserIndex, .FactionArmourEqpSlot)
                            
                            .FactionArmourEqpObjIndex = 0
                            .FactionArmourEqpSlot = 0
                        End If
                    
                        .Object(Slot).Equipped = 0
                        .ArmourEqpObjIndex = 0
                        .ArmourEqpSlot = 0
                    End With
                    
                    If .flags.Navegando <> 1 Then
                        Call DarCuerpoDesnudo(UserIndex, .flags.Mimetizado > 0)

                    End If
                    
                    .flags.Desnudo = 1 '[TEMPORAL]
                    
                    If RefreshChar Then
                        With .Char
                            Call ChangeUserChar(UserIndex, .body, .head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
                        End With
                    End If
                End If
                 
            Case eOBJType.otCASCO
                With .Invent
                    .Object(Slot).Equipped = 0
                    .CascoEqpObjIndex = 0
                    .CascoEqpSlot = 0
                End With
                
                If .flags.Mimetizado = 0 Then
                    With .Char
                        .CascoAnim = ConstantesGRH.NingunCasco
                        
                        If RefreshChar Then
                            Call ChangeUserChar(UserIndex, .body, .head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
                        End If
                    End With
                End If
            
            Case eOBJType.otESCUDO
                With .Invent
                    .Object(Slot).Equipped = 0
                    .EscudoEqpObjIndex = 0
                    .EscudoEqpSlot = 0
                End With
                
                If .flags.Mimetizado = 0 Then
                    With .Char
                        .ShieldAnim = ConstantesGRH.NingunEscudo
                        
                        If RefreshChar Then
                            Call ChangeUserChar(UserIndex, .body, .head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
                        End If
                    End With
                End If
            
            Case eOBJType.otMochilas
                With .Invent
                    .Object(Slot).Equipped = 0
                    .MochilaEqpObjIndex = 0
                    .MochilaEqpSlot = 0
                End With
                
                Call InvUsuario.TirarTodosLosItemsEnMochila(UserIndex)
                .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS
        End Select
    End With
    
    If RefreshChar Then
        Call WriteUpdateUserStats(UserIndex)
    End If
    
    Call UpdateUserInv(False, UserIndex, Slot)
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Desquipar. Error " & Err.Number & " : " & Err.Description)

End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 14/01/2010 (ZaMa)
'14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
'***************************************************

On Error GoTo ErrHandler
    
    If ObjData(ObjIndex).Mujer = 1 Then
        SexoPuedeUsarItem = UserList(UserIndex).Genero <> eGenero.Hombre
    ElseIf ObjData(ObjIndex).Hombre = 1 Then
        SexoPuedeUsarItem = UserList(UserIndex).Genero <> eGenero.Mujer
    Else
        SexoPuedeUsarItem = True
    End If
        
    Exit Function
ErrHandler:
    Call LogError("SexoPuedeUsarItem")
End Function


Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, Optional ByRef sMotivo As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 26/05/2011 (Amraphen)
'14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
'26/05/2011: Amraphen - Agrego validación para armaduras faccionarias de segunda jerarquía.
'***************************************************
On Error GoTo ErrHandler
  
Dim ArmourIndex As Integer
Dim FaltaPrimeraJerarquia As Boolean

    If ObjData(ObjIndex).Real Then
        If UserList(UserIndex).Faccion.Alignment = eCharacterAlignment.FactionRoyal And esArmada(UserIndex) Then
            If ObjData(ObjIndex).Real = 2 Then
                ArmourIndex = UserList(UserIndex).Invent.ArmourEqpObjIndex
                
                If ArmourIndex > 0 And ObjData(ArmourIndex).Real = 1 Then
                    FaccionPuedeUsarItem = True
                Else
                    FaccionPuedeUsarItem = False
                    FaltaPrimeraJerarquia = True
                End If
            Else 'Es item faccionario común
                FaccionPuedeUsarItem = True
            End If
        Else
            FaccionPuedeUsarItem = False
        End If
    ElseIf ObjData(ObjIndex).Caos Then
        If UserList(UserIndex).Faccion.Alignment = eCharacterAlignment.FactionLegion And esCaos(UserIndex) Then
            If ObjData(ObjIndex).Caos = 2 Then
                ArmourIndex = UserList(UserIndex).Invent.ArmourEqpObjIndex
                
                If ArmourIndex > 0 And ObjData(ArmourIndex).Caos = 1 Then
                    FaccionPuedeUsarItem = True
                Else
                    FaccionPuedeUsarItem = False
                    FaltaPrimeraJerarquia = True
                End If
            Else 'Es item faccionario común
                FaccionPuedeUsarItem = True
            End If
        Else
            FaccionPuedeUsarItem = False
        End If
    Else
        FaccionPuedeUsarItem = True
    End If
    
    If Not FaccionPuedeUsarItem Then
        If FaltaPrimeraJerarquia Then
            sMotivo = "Debes tener equipada una armadura faccionaria."
        Else
            sMotivo = "Tu alineación no puede usar este objeto."
        End If
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function FaccionPuedeUsarItem de InvUsuario.bas")
End Function

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
'*************************************************
'Author: Unknown
'Last modified: 03/02/2015 (D'Artagnan)
'01/08/2009: ZaMa - Now it's not sent any sound made by an invisible admin
'14/01/2010: ZaMa - Agrego el motivo especifico por el que no puede equipar/usar el item.
'26/05/2011: Amraphen - Agregadas armaduras faccionarias de segunda jerarquía.
'10/04/2012: ZaMa - Evito actualizar el cliente si no equipo/desequipo nada.
'21/02/2014: D'Artagnan - Bug fixes and code optimization.
'03/02/2015: D'Artagnan - Changes in armors behavior.
'*************************************************

On Error GoTo ErrHandler

    'Equipa un item del inventario
    Dim Obj As ObjData
    Dim ObjIndex As Integer
    Dim sMotivo As String
    
    With UserList(UserIndex)
        
        ObjIndex = .Invent.Object(Slot).ObjIndex
        Obj = ObjData(ObjIndex)
        
        If Obj.ItemGM And Not EsGm(UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "Sólo los Game Masters pueden equipar este objeto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Obj.Newbie = 1 Then
            If Not EsNewbie(UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "Sólo los newbies pueden equipar este objeto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
       
        If Not ClasePuedeUsarItem(UserIndex, ObjIndex) Then
            Call WriteConsoleMsg(UserIndex, "Tu clase no puede equipar este objeto", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not SexoPuedeUsarItem(UserIndex, ObjIndex) Then
            Call WriteConsoleMsg(UserIndex, "Tu género no puede equipar este objeto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not CheckRazaUsaRopa(UserIndex, ObjIndex) Then
            Call WriteConsoleMsg(UserIndex, "Tu raza no puede equipar este objeto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
            Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Obj.MinimumLevel > .Stats.ELV Then
            Call WriteConsoleMsg(UserIndex, "Tu nivel es muy bajo para equipar este objeto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
                        
        Select Case Obj.ObjType
            Case eOBJType.otWeapon, eOBJType.otTool
                'Si esta equipado lo quita
                If .Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot, False)
                    
                    ' Enable the berserk if it mets the requirements.
                    If HasPassiveAssigned(UserIndex, ePassiveSpells.Berserk) And Not HasPassiveActivated(UserIndex, ePassiveSpells.Berserk) Then
                        If BerzerkConditionMet(UserIndex) Then
                            Call ActivatePassive(UserIndex, ePassiveSpells.Berserk, True)
                            Call SendBerserkEffect(UserIndex, ePassiveSpells.Berserk, True)
                        End If
                    End If
                    
                    'Animacion por defecto
                    If .flags.Mimetizado Then
                        .OrigChar.WeaponAnim = ConstantesGRH.NingunArma
                    Else
                        .Char.WeaponAnim = ConstantesGRH.NingunArma
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                    
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If .Invent.WeaponEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, .Invent.WeaponEqpSlot, False)
                End If
                
                If Obj.TwoHanded = 1 And .Invent.EscudoEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, .Invent.EscudoEqpSlot, False)
                End If
                
                
                .Invent.Object(Slot).Equipped = 1
                .Invent.WeaponEqpObjIndex = ObjIndex
                .Invent.WeaponEqpSlot = Slot
                
                ' Berserk
                If HasPassiveAssigned(UserIndex, ePassiveSpells.Berserk) Then
                    Dim berserkEnabled As Boolean
                    berserkEnabled = BerzerkConditionMet(UserIndex)
                    If HasPassiveActivated(UserIndex, ePassiveSpells.Berserk) <> berserkEnabled Then
                        Call ActivatePassive(UserIndex, ePassiveSpells.Berserk, berserkEnabled)
                        Call SendBerserkEffect(UserIndex, ePassiveSpells.Berserk, berserkEnabled)
                    End If
                End If
                
                'El sonido solo se envia si no lo produce un admin invisible
                If Not (.flags.AdminInvisible = 1) Then _
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(ConstantesSonidos.SacarArma, .Pos.X, .Pos.Y, .Char.CharIndex))
                
                If .flags.Mimetizado Then
                    .OrigChar.WeaponAnim = GetWeaponAnim(UserIndex, ObjIndex)
                Else
                    .Char.WeaponAnim = GetWeaponAnim(UserIndex, ObjIndex)
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                End If
            
            Case eOBJType.otAnillo
                'Si esta equipado lo quita
                If .Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot, True)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If .Invent.AnilloEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, .Invent.AnilloEqpSlot, True)
                End If
        
                .Invent.Object(Slot).Equipped = 1
                .Invent.AnilloEqpObjIndex = ObjIndex
                .Invent.AnilloEqpSlot = Slot
            Case eOBJType.otFlechas
                'Si esta equipado lo quita
                If .Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot, True)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If .Invent.MunicionEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, .Invent.MunicionEqpSlot, True)
                End If
        
                .Invent.Object(Slot).Equipped = 1
                .Invent.MunicionEqpObjIndex = ObjIndex
                .Invent.MunicionEqpSlot = Slot
            
            Case eOBJType.otArmadura
                'Si esta equipado lo quita
                If .Invent.Object(Slot).Equipped Then
                    Call Desequipar(UserIndex, Slot, True)
                    Exit Sub
                End If
                
                'Nos fijamos si es armadura de segunda jerarquia
                If (Obj.Real = 2) Or (Obj.Caos = 2) Then
                    'Si no tiene armadura real equipada, entonces no tenemos nada que hacer
                    If .Invent.ArmourEqpObjIndex > 0 Then
                        If (ObjData(.Invent.ArmourEqpObjIndex).Real = 0) And _
                            (ObjData(.Invent.ArmourEqpObjIndex).Caos = 0) Then
                            
                            Call WriteConsoleMsg(UserIndex, "Para poder utilizar esta armadura es necesario que tengas equipada una armadura faccionaria.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                    Else
                        Call WriteConsoleMsg(UserIndex, "Para poder utilizar esta armadura es necesario que tengas equipada una armadura faccionaria.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Quita el anterior
                    If .Invent.FactionArmourEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.FactionArmourEqpSlot, True)
                    End If
                    
                    'Lo equipa
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.FactionArmourEqpObjIndex = ObjIndex
                    .Invent.FactionArmourEqpSlot = Slot
                    
                    'Si tiene la armadura necesaria para que se vea la segunda jerarquía la mostramos, sino no hacemos nada.
                    If .Invent.ArmourEqpObjIndex = ArmadurasFaccion(.Clase, .raza).Armada(eTipoDefArmors.ieMax) Or _
                       .Invent.ArmourEqpObjIndex = ArmadurasFaccion(.Clase, .raza).Armada(eTipoDefArmors.ieAlta) Or _
                       .Invent.ArmourEqpObjIndex = ArmadurasFaccion(.Clase, .raza).Caos(eTipoDefArmors.ieMax) Or _
                       .Invent.ArmourEqpObjIndex = ArmadurasFaccion(.Clase, .raza).Caos(eTipoDefArmors.ieAlta) Then
                    
                        If .flags.Mimetizado Then
                            .OrigChar.body = GetBodyForUser(UserIndex, ObjIndex)
                        Else
                            If .flags.Navegando = 0 Then
                                .Char.body = GetBodyForUser(UserIndex, ObjIndex)
                                Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                            End If
                        End If
                    End If
                Else
                    'Quita el anterior
                    If .Invent.ArmourEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.ArmourEqpSlot, True)
                    End If
            
                    'Lo equipa
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.ArmourEqpObjIndex = ObjIndex
                    .Invent.ArmourEqpSlot = Slot
                    
                    If .flags.Mimetizado Then
                        .OrigChar.body = GetBodyForUser(UserIndex, ObjIndex)
                    Else
                        If .flags.Navegando = 0 Then
                            .Char.body = GetBodyForUser(UserIndex, ObjIndex)
                            Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        End If
                    End If
                        
                    .flags.Desnudo = 0
                End If
            
            Case eOBJType.otCASCO
                'Si esta equipado lo quita
                If .Invent.Object(Slot).Equipped Then
                    Call Desequipar(UserIndex, Slot, False)
                    If .flags.Mimetizado Then
                        .OrigChar.CascoAnim = ConstantesGRH.NingunCasco
                    ElseIf .flags.Navegando = 0 Then
                        .Char.CascoAnim = ConstantesGRH.NingunCasco
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                    Exit Sub
                End If
        
                'Quita el anterior
                If .Invent.CascoEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, .Invent.CascoEqpSlot, False)
                End If
        
                'Lo equipa
                
                .Invent.Object(Slot).Equipped = 1
                .Invent.CascoEqpObjIndex = ObjIndex
                .Invent.CascoEqpSlot = Slot
                If .flags.Mimetizado Then
                    .OrigChar.CascoAnim = Obj.CascoAnim
                Else
                    .Char.CascoAnim = Obj.CascoAnim
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                End If
            
            Case eOBJType.otESCUDO
               'Si esta equipado lo quita
               If .Invent.Object(Slot).Equipped Then
                   Call Desequipar(UserIndex, Slot, False)
                   If .flags.Mimetizado Then
                       .OrigChar.ShieldAnim = ConstantesGRH.NingunEscudo
                   ElseIf .flags.Navegando = 0 Then
                       .Char.ShieldAnim = ConstantesGRH.NingunEscudo
                       Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                   End If
                    ' Enable the berserk if it mets the requirements.
                   If HasPassiveAssigned(UserIndex, ePassiveSpells.Berserk) And Not HasPassiveActivated(UserIndex, ePassiveSpells.Berserk) Then
                       If BerzerkConditionMet(UserIndex) Then
                           Call ActivatePassive(UserIndex, ePassiveSpells.Berserk, True)
                           Call SendBerserkEffect(UserIndex, ePassiveSpells.Berserk, True)
                       End If
                   End If
                   Exit Sub
                End If
        
                'Quita el anterior
                If .Invent.EscudoEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, .Invent.EscudoEqpSlot, False)
                End If
               
               If .Invent.WeaponEqpObjIndex > 0 Then
                   If ObjData(.Invent.WeaponEqpObjIndex).TwoHanded = 1 Then
                       Call Desequipar(UserIndex, .Invent.WeaponEqpSlot, False)
                   End If
               End If
               
               'Lo equipa
               .Invent.Object(Slot).Equipped = 1
               .Invent.EscudoEqpObjIndex = ObjIndex
               .Invent.EscudoEqpSlot = Slot
                
                If HasPassiveAssigned(UserIndex, ePassiveSpells.Berserk) Then
                   If HasPassiveActivated(UserIndex, ePassiveSpells.Berserk) Then
                       Call ActivatePassive(UserIndex, ePassiveSpells.Berserk, False)
                       Call SendBerserkEffect(UserIndex, ePassiveSpells.Berserk, False)
                   End If
                End If
                
                If .flags.Mimetizado Then
                    .OrigChar.ShieldAnim = Obj.ShieldAnim
                    
                ElseIf .flags.Navegando = 0 Then
                   .Char.ShieldAnim = Obj.ShieldAnim
                    
                   Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                   Call WriteUpdateUserStats(UserIndex)
                End If
                 
            Case eOBJType.otMochilas
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If .Invent.Object(Slot).Equipped Then
                    Call Desequipar(UserIndex, Slot, True)
                    Exit Sub
                End If
                If .Invent.MochilaEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, .Invent.MochilaEqpSlot, True)
                End If
                .Invent.Object(Slot).Equipped = 1
                .Invent.MochilaEqpObjIndex = ObjIndex
                .Invent.MochilaEqpSlot = Slot
                .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS + Obj.MochilaType * 5
                Call WriteAddSlots(UserIndex, Obj.MochilaType)
        End Select
    End With
    
    'Actualiza
    Call UpdateUserInv(False, UserIndex, Slot)
    
    Exit Sub
    
ErrHandler:
    Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & Err.Number & " - Error Description : " & Err.Description)
End Sub

Public Function GetBodyForUser(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
    On Error GoTo ErrHandler
    
    With UserList(UserIndex)

        Select Case .raza
            Case eRaza.Humano, eRaza.Elfo
                GetBodyForUser = IIf(.Genero = Hombre, ObjData(ObjIndex).NumRopajeHombreAlto, ObjData(ObjIndex).NumRopajeMujerAlto)
            Case eRaza.Drow
                GetBodyForUser = IIf(.Genero = Hombre, ObjData(ObjIndex).NumRopajeHombreDrow, ObjData(ObjIndex).NumRopajeMujerDrow)
            Case eRaza.Gnomo, eRaza.Enano
                GetBodyForUser = IIf(.Genero = Hombre, ObjData(ObjIndex).NumRopajeHombreBajo, ObjData(ObjIndex).NumRopajeMujerBajo)
        End Select
           
        If GetBodyForUser = 0 Then
            ' if there's no specific body defined, use the generic representation of the item
            ' if no generic representation is present, keep the current char body.
            If ObjData(ObjIndex).NumRopajeGenerico > 0 Then
            GetBodyForUser = ObjData(ObjIndex).NumRopajeGenerico
            Else
                GetBodyForUser = .Char.body
        End If

            Exit Function
        End If
        
    End With
        
    Exit Function
ErrHandler:
    Call LogError("Error GetBodyForUser ItemIndex:" & ObjIndex & " and UserIndex: " & UserIndex)
    
End Function

Public Function CheckRazaUsaRopa(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 14/01/2010 (ZaMa)
'14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)
        If ObjData(ItemIndex).RazaProhibida(1) <> 0 Then
            Dim I As Integer
            For I = 1 To NUMRAZAS
                If ObjData(ItemIndex).RazaProhibida(I) = .raza Then
                    CheckRazaUsaRopa = False
                    Exit Function
                End If
            Next I
        End If
        
        CheckRazaUsaRopa = True
    End With
    
    Exit Function
    
ErrHandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
'*************************************************
'Author: Unknown
'Last modified: 08/11/2015
'Handels the usage of items from inventory box.
'24/01/2007 Pablo (ToxicWaste) - Agrego el Cuerno de la Armada y la Legión.
'24/01/2007 Pablo (ToxicWaste) - Utilización nueva de Barco en lvl 20 por clase Pirata y Pescador.
'01/08/2009: ZaMa - Now it's not sent any sound made by an invisible admin, except to its own client
'17/11/2009: ZaMa - Ahora se envia una orientacion de la posicion hacia donde esta el que uso el cuerno.
'27/11/2009: Budi - Se envia indivualmente cuando se modifica a la Agilidad o la Fuerza del personaje.
'08/12/2009: ZaMa - Agrego el uso de hacha de madera elfica.
'10/12/2009: ZaMa - Arreglos y validaciones en todos las herramientas de trabajo.
'08/11/2015: D'Artagnan - Boat usage: added class restrictions.
'*************************************************
On Error GoTo ErrHandler
  

    Dim Obj As ObjData
    Dim ObjIndex As Integer
    Dim TargObj As ObjData
    Dim MiObj As Obj
    
    With UserList(UserIndex)
    
        If .Invent.Object(Slot).Amount = 0 Then Exit Sub
        
        Obj = ObjData(.Invent.Object(Slot).ObjIndex)
        
        If .flags.Petrificado Then
            Call WriteConsoleMsg(UserIndex, "¡Estás petrificado!.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Obj.Newbie = 1 Then
            If Not EsNewbie(UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "Sólo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        If Obj.ObjType = eOBJType.otWeapon Then
            If Obj.proyectil = 1 Then
                'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
                If Not IntervaloPermiteUsar(UserIndex, False) Then Exit Sub
            Else
                'dagas
                If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
            End If
        Else
            If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
        End If
        
        If Obj.MinimumLevel > .Stats.ELV Then
            Call WriteConsoleMsg(UserIndex, "Tu nivel es muy bajo para usar este objeto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        ObjIndex = .Invent.Object(Slot).ObjIndex
        .flags.TargetObjInvIndex = ObjIndex
        .flags.TargetObjInvSlot = Slot
        
        Select Case Obj.ObjType
            Case eOBJType.otUseOnce
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
        
                'Usa el item
                .Stats.MinHam = .Stats.MinHam + Obj.MinHam
                If .Stats.MinHam > .Stats.MaxHam Then _
                    .Stats.MinHam = .Stats.MaxHam
                .flags.Hambre = 0
                Call WriteUpdateHungerAndThirst(UserIndex)
                'Sonido
                
                If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
                    Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MORFAR_MANZANA)
                Else
                    Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.SOUND_COMIDA)
                End If
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
        
            Case eOBJType.otGuita
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                .Stats.GLD = .Stats.GLD + .Invent.Object(Slot).Amount
                .Invent.Object(Slot).Amount = 0
                .Invent.Object(Slot).ObjIndex = 0
                .Invent.NroItems = .Invent.NroItems - 1
                
                Call UpdateUserInv(False, UserIndex, Slot)
                Call WriteUpdateGold(UserIndex)
                
            Case eOBJType.otWeapon, eOBJType.otTool
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If Not .Stats.MinSta > 0 Then
                    Call WriteConsoleMsg(UserIndex, "Estás muy cansad" & _
                                IIf(.Genero = eGenero.Hombre, "o", "a") & ".", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If ObjData(ObjIndex).proyectil = 1 Then
                    If .Invent.Object(Slot).Equipped = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberías equipartela.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Proyectiles)  'Call WriteWorkRequestTarget(UserIndex, Proyectiles)
                Else
                    If Obj.ObjType = eOBJType.otTool Then
                        If .Invent.WeaponEqpObjIndex = ObjIndex Then
                            Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, Professions(Obj.ProfessionType).SkillNumber)
                        Else
                             Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                        End If
                        Exit Sub
                    End If
                    
                    Select Case ObjIndex
                    
                        Case ConstantesItems.CañaPesca, ConstantesItems.RedPesca, ConstantesItems.CañaPescaNW
                            
                            ' Lo tiene equipado?
                            If .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Pesca)  'Call WriteWorkRequestTarget(UserIndex, eSkill.Pesca)
                            Else
                                 Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                            End If
                                            
                        Case ConstantesItems.PiqueteMinero, ConstantesItems.PiqueteMineroNW
                        
                            ' Lo tiene equipado?
                            If .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Mineria)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                            End If
                                                     
                        Case Else ' Every other tool should be using the same crafting system, based on how the tool and the professions are configured.
                        
                            If .Invent.WeaponEqpObjIndex <> ObjIndex Then
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            Dim ProfessionType As Byte
                            ProfessionType = Obj.ProfessionType
                            
                            If Obj.ProfessionType <= 0 Or Obj.ProfessionType > UBound(Professions) Then Exit Sub
                            If Not Professions(Obj.ProfessionType).Enabled Then
                                Call WriteConsoleMsg(UserIndex, "La profesión se encuentra desactivada. Intenta denuevo más tarde.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            Call WriteCraftableRecipes(UserIndex, Obj.ProfessionType)
                            Call WriteShowCraftForm(UserIndex)
                    End Select
                End If
            
            Case eOBJType.otPociones
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If Not IntervaloPermiteGolpeUsar(UserIndex, False) Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Debes esperar unos momentos para tomar otra poción!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                Dim PotionConsumed As Boolean
                PotionConsumed = modSystem_ObjectActions.Potions_Use(UserIndex, Obj)
                
                If PotionConsumed Then _
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                
                Select Case .flags.TipoPocion
                
                    Case 5 ' Pocion violeta
                        If .flags.Envenenado = 1 Then
                            .flags.Envenenado = 0
                            Call WriteConsoleMsg(UserIndex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)
                        End If
                        'Quitamos del inv el item
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call SendData(ToUser, UserIndex, PrepareMessagePlayWave(ConstantesSonidos.Tomar, .Pos.X, .Pos.Y, .Char.CharIndex))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(ConstantesSonidos.Tomar, .Pos.X, .Pos.Y, .Char.CharIndex))
                        End If
                        
                    Case 6  ' Pocion Negra
                        If .flags.Privilegios And PlayerType.User Then
                        
                            If .flags.DueloIndex > 0 Then
                                Call WriteConsoleMsg(UserIndex, "No puedes usar este objeto durante un duelo.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                        
                            Call QuitarUserInvItem(UserIndex, Slot, 1)
                            Call UserDie(UserIndex)
                            Call WriteConsoleMsg(UserIndex, "Sientes un gran mareo y pierdes el conocimiento.", FontTypeNames.FONTTYPE_FIGHT)
                        End If
               End Select
               
               Call WriteUpdateUserStats(UserIndex)
        
             Case eOBJType.otBebidas
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed
                If .Stats.MinAGU > .Stats.MaxAGU Then _
                    .Stats.MinAGU = .Stats.MaxAGU
                .flags.Sed = 0
                Call WriteUpdateHungerAndThirst(UserIndex)
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call SendData(ToUser, UserIndex, PrepareMessagePlayWave(ConstantesSonidos.Tomar, .Pos.X, .Pos.Y, .Char.CharIndex))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(ConstantesSonidos.Tomar, .Pos.X, .Pos.Y, .Char.CharIndex))
                End If
            
            Case eOBJType.otLlaves
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If .flags.TargetObj = 0 Then Exit Sub
                TargObj = ObjData(.flags.TargetObj)
                '¿El objeto clickeado es una puerta?
                If TargObj.ObjType = eOBJType.otPuertas Then
                    '¿Esta cerrada?
                    If TargObj.Cerrada = 1 Then
                          '¿Cerrada con llave?
                          If TargObj.Llave > 0 Then
                             If TargObj.clave = Obj.clave Then
                 
                                MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex _
                                = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
                                .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                                Call WriteConsoleMsg(UserIndex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                             Else
                                Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                             End If
                          Else
                             If TargObj.clave = Obj.clave Then
                                MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex _
                                = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
                                Call WriteConsoleMsg(UserIndex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
                                .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                                Exit Sub
                             Else
                                Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                             End If
                          End If
                    Else
                          Call WriteConsoleMsg(UserIndex, "No está cerrada.", FontTypeNames.FONTTYPE_INFO)
                          Exit Sub
                    End If
                End If
            
            Case eOBJType.otBotellaVacia
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                Dim clickPos As WorldPos
                clickPos.Map = .flags.TargetMap
                clickPos.X = .flags.TargetX
                clickPos.Y = .flags.TargetY
                
                If Not HayAgua(.Pos.Map, .flags.TargetX, .flags.TargetY) Then
                    Call WriteConsoleMsg(UserIndex, "No hay agua allí.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If Distancia(clickPos, .Pos) > 2 Then
                    Call WriteConsoleMsg(UserIndex, "Estás muy lejos para realizar esta acción.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                MiObj.Amount = 1
                MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexAbierta
                
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(.Pos, MiObj)
                End If
            
            Case eOBJType.otBotellaLlena
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed
                If .Stats.MinAGU > .Stats.MaxAGU Then _
                    .Stats.MinAGU = .Stats.MaxAGU
                    
                .flags.Sed = 0
                
                Call WriteUpdateHungerAndThirst(UserIndex)
                
                MiObj.Amount = 1
                MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexCerrada
                
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(.Pos, MiObj)
                End If
            
            Case eOBJType.otPergaminos
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If .Stats.MaxMan > 0 Then
                    If .flags.Hambre = 0 And _
                        .flags.Sed = 0 Then
                        Call AgregarHechizo(UserIndex, Slot)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Estás demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_INFO)
                End If
                
            Case eOBJType.otMinerales
                If .flags.Muerto = 1 Then
                     Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub
                End If
                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, FundirMetal) 'Call WriteWorkRequestTarget(UserIndex, FundirMetal)
               
            Case eOBJType.otInstrumentos
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If Obj.Real Then '¿Es el Cuerno Real?
                    If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                        If MapInfo(.Pos.Map).Pk = False Then
                            Call WriteConsoleMsg(UserIndex, "No hay peligro aquí. Es zona segura.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call SendData(ToUser, UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y, .Char.CharIndex))
                        Else
                            Call AlertarFaccionarios(UserIndex)
                            Call SendData(SendTarget.toMap, .Pos.Map, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y, .Char.CharIndex))
                        End If
                        
                        Exit Sub
                    Else
                        Call WriteConsoleMsg(UserIndex, "Sólo miembros del ejército real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                ElseIf Obj.Caos Then '¿Es el Cuerno Legión?
                    If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                        If MapInfo(.Pos.Map).Pk = False Then
                            Call WriteConsoleMsg(UserIndex, "No hay peligro aquí. Es zona segura.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call SendData(ToUser, UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y, .Char.CharIndex))
                        Else
                            Call AlertarFaccionarios(UserIndex)
                            Call SendData(SendTarget.toMap, .Pos.Map, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y, .Char.CharIndex))
                        End If
                        
                        Exit Sub
                    Else
                        Call WriteConsoleMsg(UserIndex, "Sólo miembros de la legión oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                End If
                'Si llega aca es porque es o Laud o Tambor o Flauta
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call SendData(ToUser, UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y, .Char.CharIndex))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y, .Char.CharIndex))
                End If
               
            Case eOBJType.otBarcos
                'Verifica si esta aproximado al agua antes de permitirle navegar
                Dim sMessage As String
                
                If ClasePuedeUsarItem(UserIndex, ObjIndex) Then
                    If EsGm(UserIndex) Or _
                        (UserAreaHasWater(UserIndex) And .flags.Navegando = 0) Or _
                        (UserAreaHasLand(UserIndex) And .flags.Navegando = 1) Then
                        Call DoNavega(UserIndex, Slot)
                    Else
                        Call WriteConsoleMsg(UserIndex, "¡Debes aproximarte a la costa para usar el barco!", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto", FontTypeNames.FONTTYPE_INFO)
                End If
    
            Case eOBJType.otTrigger
                ' Dead?
                If .flags.Muerto = 1 Then
                     Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub
                End If
                
                ' Can use it?
                Dim Motivo As String
                If Not ClasePuedeUsarItem(UserIndex, ObjIndex) Then
                    Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                ' Set it
                Call DropObj(UserIndex, Slot, 1, .Pos.Map, .Pos.X, .Pos.Y)
                
            Case eOBJType.otSurpriseBox
                ' Dead?
                If .flags.Muerto = 1 Then
                     Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub
                End If
                
                MiObj = SurpriseObj(ObjIndex)
                
                If MiObj.ObjIndex <> 0 Then
                    Call WriteConsoleMsg(UserIndex, "¡Has obtenido " & ObjData(MiObj.ObjIndex).Name & "(" & MiObj.Amount & ") de la caja sorpresa!", FontTypeNames.FONTTYPE_INFO)
                    
                    If Not MeterItemEnInventario(UserIndex, MiObj) Then
                        Call TirarItemAlPiso(.Pos, MiObj)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Parece que la caja sorpresa no ha arrojado ningún objeto.", FontTypeNames.FONTTYPE_INFO)
                End If
                
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
            Case eOBJType.otGuildBook
                ' Dead?
                If .flags.Muerto = 1 Then
                     Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub
                End If
                
                If .Guild.IdGuild = 0 Then
                    Call WriteConsoleMsg(UserIndex, "No perteneces a ningún clan.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                Call WriteConsoleMsg(UserIndex, "¡Hás aumentado el cupo máximo de miembros de tu clan en " & ObjData(ObjIndex).Cupos & "!.", FontTypeNames.FONTTYPE_INFO)
                
                Call QuitarUserInvItem(UserIndex, Slot, 1)
        End Select
    
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UseInvItem de InvUsuario.bas")
End Sub

Private Function SurpriseObj(ByVal ObjIndex As Integer) As Obj
'***************************************************
'Author: Pato
'Last Modification: 12/29/2012
'
'***************************************************
On Error GoTo ErrHandler
  
Dim Obj As Obj
Dim Num As Long
Dim I As Long

'num = RandomNumber(1, MAX_DIGIT_SB_RND)


With ObjData(ObjIndex).SurpriseDrops
    Num = RandomNumber(1, .NroItems)
    
    If .Drop(Num).ObjIndex <> 0 Then
        Obj.ObjIndex = .Drop(Num).ObjIndex
        Obj.Amount = .Drop(Num).Amount
    End If
    
    'Do While (I < .NroItems)
    '    I = I + 1
    '
    '    If num <= .Drop(I).prob Then
    '        Obj.ObjIndex = .Drop(I).ObjIndex
    '        Obj.Amount = .Drop(I).Amount
    '        Exit Do
    '    End If
    'Loop
End With

SurpriseObj = Obj
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function SurpriseObj de InvUsuario.bas")
End Function

Sub TirarTodo(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = 6 Then Exit Sub
        
        Call TirarTodosLosItems(UserIndex)
        
    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en TirarTodo. Error: " & Err.Number & " - " & Err.Description)
End Sub

Public Function ItemSeCae(ByVal Index As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler

    With ObjData(Index)
        ItemSeCae = (.Real <> 1 Or .NoSeCae = 0) And _
                    (.Caos <> 1 Or .NoSeCae = 0) And _
                    .ObjType <> eOBJType.otLlaves And _
                    .ObjType <> eOBJType.otBarcos And _
                    .ObjType <> otQuest And _
                    .NoSeCae = 0 And _
                    .Intransferible = 0
    End With

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ItemSeCae de InvUsuario.bas")
End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2010 (ZaMa)
'12/01/2010: ZaMa - Ahora los piratas no explotan items solo si estan entre 20 y 25
'22b/08/2015: Nightw - Workers don't drop an equiped fishing net if they are navigating
'***************************************************
On Error GoTo ErrHandler

    Dim I As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    Dim DropAgua As Boolean
    Dim canDrop As Boolean
    canDrop = True
    With UserList(UserIndex)
        For I = 1 To .CurrentInventorySlots
            canDrop = True
            ItemIndex = .Invent.Object(I).ObjIndex
            If ItemIndex > 0 Then
                 If ItemSeCae(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    
                    'Creo el Obj
                    MiObj.Amount = .Invent.Object(I).Amount
                    MiObj.ObjIndex = ItemIndex

                    DropAgua = True
                    
                    Select Case .clase
                        Case eClass.Thief
                            ' Si tiene galeon equipado
                            If .Invent.BarcoObjIndex = 476 Then
                                ' Limitación por nivel, después dropea normalmente
                                If .Stats.ELV = 28 Then
                                    ' No dropea en agua
                                    DropAgua = False
                                End If
                            End If
                        
                        Case eClass.Worker
                            If .flags.Navegando And .Invent.Object(I).Equipped And ItemIndex = RED_PESCA Then
                                canDrop = False
                            End If
                    End Select
                   
                    
                    Call Tilelibre(.Pos, NuevaPos, MiObj, DropAgua, True)
                    
                    If canDrop And NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(UserIndex, I, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                    End If
                 End If
            End If
        Next I
    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error en TirarTodosLosItems. Error: " & Err.Number & " - " & Err.Description)
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
    
    ItemNewbie = ObjData(ItemIndex).Newbie = 1
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ItemNewbie de InvUsuario.bas")
End Function

Sub TirarTodosLosItemsNoNewbies(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 23/11/2009
'07/11/09: Pato - Fix bug #2819911
'23/11/2009: ZaMa - Optimizacion de codigo.
'***************************************************
On Error GoTo ErrHandler
  
    Dim I As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    
    With UserList(UserIndex)
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = 6 Then Exit Sub
            
        Dim Cantidad As Long
        Cantidad = .Stats.GLD - CLng(.Stats.ELV) * 10000
        
        If Cantidad > 0 Then _
            Call TirarOro(Cantidad, UserIndex)
            
        For I = 1 To UserList(UserIndex).CurrentInventorySlots
            ItemIndex = .Invent.Object(I).ObjIndex
            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    
                    'Creo MiObj
                    MiObj.Amount = .Invent.Object(I).Amount
                    MiObj.ObjIndex = ItemIndex
                    'Pablo (ToxicWaste) 24/01/2007
                    'Tira los Items no newbies en todos lados.
                    Tilelibre .Pos, NuevaPos, MiObj, True, True
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(UserIndex, I, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                    End If
                End If
            End If
        Next I
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub TirarTodosLosItemsNoNewbies de InvUsuario.bas")
End Sub

Sub TirarTodosLosItemsEnMochila(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/09 (Budi)
'***************************************************
On Error GoTo ErrHandler
  
    Dim I As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    
    With UserList(UserIndex)
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = 6 Then Exit Sub
        
        For I = MAX_NORMAL_INVENTORY_SLOTS + 1 To .CurrentInventorySlots
            ItemIndex = .Invent.Object(I).ObjIndex
            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    
                    'Creo MiObj
                    MiObj.Amount = .Invent.Object(I).Amount
                    MiObj.ObjIndex = ItemIndex
                    Tilelibre .Pos, NuevaPos, MiObj, True, True
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(UserIndex, I, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                    End If
                End If
            End If
        Next I
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub TirarTodosLosItemsEnMochila de InvUsuario.bas")
End Sub

Public Sub moveItem(ByVal UserIndex As Integer, ByVal originalSlot As Integer, ByVal newSlot As Integer)

'**************************************************************
'Autor: -
'Last Modification: 04/04/2017
'30/03/2017: G Toyz - Retabulado, ahora los objetos se unen en caso de ser el mismo y tener una cantidad menor a 10.000.
'04/04/2017: G Toyz - Ahora sólo se switchean los objetos que no son del mismo tipo. Los demás se tratan de unir y en caso _
                                      de superar la cantidad 10.000 en objetos entre ambos lo que hace es dejar en 10000 al lugar en donde _
                                      draggeaste y la posición original del objeto que quisiste mover queda el resto de 10000.
'**************************************************************

On Error GoTo ErrHandler
  
    Dim tmpObj As UserOBJ
    Dim newObjIndex As Integer
    Dim originalObjIndex As Integer
    Dim cantidadTotal As Integer
    Dim unionObjetos As Boolean
    Dim cantidadFaltanteNuevoSlot As Integer '
    
    If (originalSlot <= 0) Or (newSlot <= 0) Then Exit Sub

    With UserList(UserIndex)
     
        If (originalSlot > .CurrentInventorySlots) Or (newSlot > .CurrentInventorySlots) Then Exit Sub
        
        'G Toyz:
        
        If .Invent.Object(originalSlot).ObjIndex = .Invent.Object(newSlot).ObjIndex Then
            cantidadTotal = .Invent.Object(newSlot).Amount + .Invent.Object(originalSlot).Amount
            If cantidadTotal <= MAX_INVENTORY_OBJS Then
                .Invent.Object(newSlot).Amount = cantidadTotal
                If .Invent.Object(originalSlot).Equipped = 1 Then
                    .Invent.Object(newSlot).Equipped = 1
                    .Invent.Object(originalSlot).Equipped = 0
                End If
                .Invent.Object(originalSlot).Amount = 0
                .Invent.Object(originalSlot).ObjIndex = 0
                unionObjetos = True
            Else
                cantidadFaltanteNuevoSlot = MAX_INVENTORY_OBJS - .Invent.Object(newSlot).Amount
                If cantidadFaltanteNuevoSlot > 0 Then
                    .Invent.Object(originalSlot).Amount = .Invent.Object(originalSlot).Amount - cantidadFaltanteNuevoSlot
                    .Invent.Object(newSlot).Amount = .Invent.Object(newSlot).Amount + cantidadFaltanteNuevoSlot
                    unionObjetos = True
                End If
            End If
        End If
        
        If unionObjetos = False Then
            tmpObj = .Invent.Object(originalSlot)
            .Invent.Object(originalSlot) = .Invent.Object(newSlot)
            .Invent.Object(newSlot) = tmpObj
        End If
        
        '//G Toyz
    
    'Viva VB6 y sus putas deficiencias.
    If .Invent.AnilloEqpSlot = originalSlot Then
        .Invent.AnilloEqpSlot = newSlot
    ElseIf .Invent.AnilloEqpSlot = newSlot Then
        .Invent.AnilloEqpSlot = originalSlot
    End If
    
    If .Invent.ArmourEqpSlot = originalSlot Then
        .Invent.ArmourEqpSlot = newSlot
    ElseIf .Invent.ArmourEqpSlot = newSlot Then
        .Invent.ArmourEqpSlot = originalSlot
    End If
    
    If .Invent.BarcoSlot = originalSlot Then
        .Invent.BarcoSlot = newSlot
    ElseIf .Invent.BarcoSlot = newSlot Then
        .Invent.BarcoSlot = originalSlot
    End If
    
    If .Invent.CascoEqpSlot = originalSlot Then
         .Invent.CascoEqpSlot = newSlot
    ElseIf .Invent.CascoEqpSlot = newSlot Then
         .Invent.CascoEqpSlot = originalSlot
    End If
    
    If .Invent.EscudoEqpSlot = originalSlot Then
        .Invent.EscudoEqpSlot = newSlot
    ElseIf .Invent.EscudoEqpSlot = newSlot Then
        .Invent.EscudoEqpSlot = originalSlot
    End If
    
    If .Invent.MochilaEqpSlot = originalSlot Then
        .Invent.MochilaEqpSlot = newSlot
    ElseIf .Invent.MochilaEqpSlot = newSlot Then
        .Invent.MochilaEqpSlot = originalSlot
    End If
    
    If .Invent.MunicionEqpSlot = originalSlot Then
        .Invent.MunicionEqpSlot = newSlot
    ElseIf .Invent.MunicionEqpSlot = newSlot Then
        .Invent.MunicionEqpSlot = originalSlot
    End If
    
    If .Invent.WeaponEqpSlot = originalSlot Then
        .Invent.WeaponEqpSlot = newSlot
    ElseIf .Invent.WeaponEqpSlot = newSlot Then
        .Invent.WeaponEqpSlot = originalSlot
    End If

    Call UpdateUserInv(False, UserIndex, originalSlot)
    Call UpdateUserInv(False, UserIndex, newSlot)
End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub moveItem de InvUsuario.bas")
End Sub

Public Sub AddInventoryItem(ByVal nUserIndex As Integer, ByVal nItemID As Integer, ByVal nAmount As Integer, _
                            Optional ByVal bPull As Boolean = True)
'******************************************
'Author: D'Artagnan
'Date: 24/01/2015
'If possible, add the specified item to the inventory.
'Otherwise, pull the item to the floor if bPull is True.
'******************************************
On Error GoTo ErrHandler
  
    Dim Item As Obj
    
    Item.ObjIndex = nItemID
    Item.Amount = nAmount
    
    If Not MeterItemEnInventario(nUserIndex, Item) And bPull Then
        Call TirarItemAlPiso(UserList(nUserIndex).Pos, Item)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AddInventoryItem de InvUsuario.bas")
End Sub

Public Function IsSecondaryArmour(ByVal nObjectIndex As Integer) As Boolean
'***************************************************
'Author: D'Artagnan
'Last Modification: 03/02/2015
'Return True if the specified object index belongs
'to a secondary armour. False otherwise.
'***************************************************
On Error GoTo ErrHandler
  
    IsSecondaryArmour = ObjData(nObjectIndex).Real = 2 Or _
                        ObjData(nObjectIndex).Caos = 2
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function IsSecondaryArmour de InvUsuario.bas")
End Function

Public Function DropAllowed(ByVal nUserIndex As Integer, Optional ByVal bLogCheating As Boolean = True) As Boolean
'***************************************************
'Author: D'Artagnan
'Last Modification: 01/03/2015
'Return True if the specified user can drop items.
'If the user can drop items, bLogCheating also
'determines if the connection will be closed.
'***************************************************
On Error GoTo ErrHandler
  
    DropAllowed = True
    
    If isTrading(nUserIndex) Then
        If isTradingWithUser(nUserIndex) Then
            ' Might be waiting for an answer.
            DropAllowed = Not (UserList(getTradingUser(nUserIndex)).flags.Comerciando = Not nUserIndex)
        Else
            ' Trading with a NPC.
            DropAllowed = False
        End If
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function DropAllowed de InvUsuario.bas")
End Function
