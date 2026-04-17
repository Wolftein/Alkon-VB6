Attribute VB_Name = "InvNpc"
'Argentum Online 0.14.0
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
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Inv & Obj
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Modulo para controlar los objetos y los inventarios.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
Public Function TirarItemAlPiso(Pos As WorldPos, Obj As Obj, Optional NotPirata As Boolean = True) As WorldPos
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim NuevaPos As WorldPos
    NuevaPos.X = 0
    NuevaPos.Y = 0
    
    Tilelibre Pos, NuevaPos, Obj, NotPirata, True
    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
        Call MakeObj(Obj, Pos.Map, NuevaPos.X, NuevaPos.Y)
    End If
    TirarItemAlPiso = NuevaPos

    Exit Function
ErrHandler:

End Function

Public Sub NPCDropArrows(ByRef npc As npc, ByVal nUserIndex As Integer)
'***************************************************
'Author: D'Artagnan (original version implemented by Mithrandir)
'Last Modification: 08/11/2015
'07/01/2014: Mithrandir - Drop arrow objects.
'08/11/2015: D'Artagnan - Earn skill experience.
'***************************************************
On Error GoTo ErrHandler
  
    Dim MiObj As Obj
    Dim bDrop As Boolean
    Dim I As Long
     
    'Flecha en NPC
    For I = 1 To 6
        With npc
            If .TengoFlechas(I) > 0 Then
                'modificar más adelante
                MiObj.Amount = Porcentaje(RandomNumber(10, 30), .TengoFlechas(I))
                
                If MiObj.Amount <> 0 Then
                    Select Case I
                        Case 1
                            MiObj.ObjIndex = ConstantesItems.Flecha      ' Flecha
                        Case 2
                            MiObj.ObjIndex = ConstantesItems.Flecha1  ' Fecha + 1
                        Case 3
                            MiObj.ObjIndex = ConstantesItems.Flecha2  ' Flecha +2
                        Case 4
                            MiObj.ObjIndex = ConstantesItems.Flecha3 ' Flecha +3
                        Case 5
                            MiObj.ObjIndex = ConstantesItems.FlechaNewbie ' Flecha Newbie
                        Case 6
                            MiObj.ObjIndex = ConstantesItems.Cuchillas 'Cuchillas
                    End Select
                    
                    Call TirarItemAlPiso(.Pos, MiObj)
                    bDrop = True
                End If
            End If
        End With
    Next
    
    If bDrop Then
        Call SubirSkill(nUserIndex, eSkill.Supervivencia, True)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub NPCDropArrows de Modulo_InventANDobj.bas")
End Sub

Public Sub NPC_TIRAR_ITEMS(ByRef npc As npc, ByVal IsPretoriano As Boolean, ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 28/11/2009
'Give away npc's items.
'28/11/2009: ZaMa - Implementado drops complejos
'02/04/2010: ZaMa - Los pretos vuelven a tirar oro.
'10/04/2011: ZaMa - Logueo los objetos logueables dropeados.
'***************************************************
On Error GoTo ErrHandler
  
On Error Resume Next

    With npc
        Dim I As Byte
        Dim a As Byte
        Dim MiObj As Obj
        Dim NroDrop As Integer
        Dim Random As Integer
        Dim ObjIndex As Integer
        Dim canDropGuildItem As Boolean
        Dim shouldLevelExpMod As Boolean
        
        canDropGuildItem = UserList(UserIndex).Guild.IdGuild <= 0
        ' This property will be set based on the difference between the user level
        ' and the NPC level. This, momentarly, will apply only to the drop of gold.
        shouldLevelExpMod = ShouldApplyExpMod(UserIndex, NpcIndex)
        
        ' Si es pretoriano, dropea el oro configurado en GiveGLD
        If IsPretoriano Then
            ' Dropea oro?
            If .GiveGLD > 0 Then _
                Call TirarOroNpc(.GiveGLD, .Pos)
        End If
        
        For I = 1 To .NroDrops
            If RandomDecimalNumber(0, 100) <= .Drop(I).Probabilidad Then
                For a = 1 To DropData(.Drop(I).DropIndex).NumItems
                    ObjIndex = DropData(.Drop(I).DropIndex).Item(a).ObjIndex
                    If ObjIndex > 0 Then
                        'Exit if we can't drop guild items.
                        If ObjIndex = ConstantesItems.RequiredGuildItem And Not canDropGuildItem Then
                            Exit For
                        End If
                    
                    
                        If ObjIndex = ConstantesItems.Oro Then
                            ' Only drop gold if the level difference allows it.
                            If Not shouldLevelExpMod Then
                                Call TirarOroNpc(DropData(.Drop(I).DropIndex).Item(a).Amount, npc.Pos)
                            End If
                        Else
                            ' Only drop items if the level mod don't apply, or if it apply and the item is a crafting material
                            If Not shouldLevelExpMod Or (shouldLevelExpMod And ObjData(ObjIndex).ObjType = otCraftingMaterial) Then
                                MiObj.Amount = DropData(.Drop(I).DropIndex).Item(a).Amount
                                MiObj.ObjIndex = ObjIndex
                                
                                Call TirarItemAlPiso(.Pos, MiObj)
                                
                                If ObjData(ObjIndex).Log = 1 Then
                                    Call LogDesarrollo(npc.Name & " dropeó " & MiObj.Amount & " " & _
                                        ObjData(ObjIndex).Name & "[" & ObjIndex & "]")
                                End If
                            End If
                        End If
                    End If
                Next a
                If .Drop(I).NoExcluyente = 0 Then Exit For
            End If
        Next I
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub NPC_TIRAR_ITEMS de Modulo_InventANDobj.bas")
End Sub

Function QuedanItems(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

On Error Resume Next

    Dim I As Integer
    If Npclist(NpcIndex).Invent.NroItems > 0 Then
        For I = 1 To MAX_INVENTORY_SLOTS
            If Npclist(NpcIndex).Invent.Object(I).ObjIndex = ObjIndex Then
                QuedanItems = True
                Exit Function
            End If
        Next
    End If
    QuedanItems = False
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function QuedanItems de Modulo_InventANDobj.bas")
End Function

''
' Gets the amount of a certain item that an npc has.
'
' @param npcIndex Specifies reference to npcmerchant
' @param ObjIndex Specifies reference to object
' @return   The amount of the item that the npc has
' @remarks This function reads the Npc.dat file
Function EncontrarCant(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Integer
'***************************************************
'Author: Unknown
'Last Modification: 03/09/08
'Last Modification By: Marco Vanotti (Marco)
' - 03/09/08 EncontrarCant now returns 0 if the npc doesn't have it (Marco)
'***************************************************
On Error GoTo ErrHandler
  
On Error Resume Next
'Devuelve la cantidad original del obj de un npc

    Dim ln As String, npcfile As String
    Dim I As Integer
    
    npcfile = DatPath & "NPCs.dat"
     
    For I = 1 To MAX_INVENTORY_SLOTS
        ln = GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "Obj" & I)
        If ObjIndex = Val(ReadField(1, ln, 45)) Then
            EncontrarCant = Val(ReadField(2, ln, 45))
            Exit Function
        End If
    Next
                       
    EncontrarCant = 0

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function EncontrarCant de Modulo_InventANDobj.bas")
End Function

Sub ResetNpcInv(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

On Error Resume Next

    Dim I As Integer
    
    With Npclist(NpcIndex)
        .Invent.NroItems = 0
        
        For I = 1 To MAX_INVENTORY_SLOTS
           .Invent.Object(I).ObjIndex = 0
           .Invent.Object(I).Amount = 0
        Next I
        
        .InvReSpawn = 0
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetNpcInv de Modulo_InventANDobj.bas")
End Sub

''
' Removes a certain amount of items from a slot of an npc's inventory
'
' @param npcIndex Specifies reference to npcmerchant
' @param Slot Specifies reference to npc's inventory's slot
' @param antidad Specifies amount of items that will be removed
Sub QuitarNpcInvItem(ByVal NpcIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 23/11/2009
'Last Modification By: Marco Vanotti (Marco)
' - 03/09/08 Now this sub checks that te npc has an item before respawning it (Marco)
'23/11/2009: ZaMa - Optimizacion de codigo.
'***************************************************
On Error GoTo ErrHandler
  
    Dim ObjIndex As Integer
    Dim iCant As Integer
    
    With Npclist(NpcIndex)
        ObjIndex = .Invent.Object(Slot).ObjIndex
    
        'Quita un Obj
        If ObjData(.Invent.Object(Slot).ObjIndex).Crucial = 0 Then
            .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - Cantidad
            
            If .Invent.Object(Slot).Amount <= 0 Then
                .Invent.NroItems = .Invent.NroItems - 1
                .Invent.Object(Slot).ObjIndex = 0
                .Invent.Object(Slot).Amount = 0
                If .Invent.NroItems = 0 And .InvReSpawn <> 1 Then
                   Call CargarInvent(NpcIndex) 'Reponemos el inventario
                End If
            End If
        Else
            .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - Cantidad
            
            If .Invent.Object(Slot).Amount <= 0 Then
                .Invent.NroItems = .Invent.NroItems - 1
                .Invent.Object(Slot).ObjIndex = 0
                .Invent.Object(Slot).Amount = 0
                
                If Not QuedanItems(NpcIndex, ObjIndex) Then
                    'Check if the item is in the npc's dat.
                    iCant = EncontrarCant(NpcIndex, ObjIndex)
                    If iCant Then
                        .Invent.Object(Slot).ObjIndex = ObjIndex
                        .Invent.Object(Slot).Amount = iCant
                        .Invent.NroItems = .Invent.NroItems + 1
                    End If
                End If
                
                If .Invent.NroItems = 0 And .InvReSpawn <> 1 Then
                   Call CargarInvent(NpcIndex) 'Reponemos el inventario
                End If
            End If
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub QuitarNpcInvItem de Modulo_InventANDobj.bas")
End Sub

Sub CargarInvent(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    'Vuelve a cargar el inventario del npc NpcIndex
    Dim LoopC As Integer
    Dim ln As String
    Dim npcfile As String
    
    npcfile = DatPath & "NPCs.dat"
    
    With Npclist(NpcIndex)
        .Invent.NroItems = Val(GetVar(npcfile, "NPC" & .Numero, "NROITEMS"))
        
        For LoopC = 1 To .Invent.NroItems
            ln = GetVar(npcfile, "NPC" & .Numero, "Obj" & LoopC)
            .Invent.Object(LoopC).ObjIndex = Val(ReadField(1, ln, 45))
            .Invent.Object(LoopC).Amount = Val(ReadField(2, ln, 45))
            
        Next LoopC
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CargarInvent de Modulo_InventANDobj.bas")
End Sub


Public Sub TirarOroNpc(ByVal Cantidad As Long, ByRef Pos As WorldPos)
'***************************************************
'Autor: ZaMa
'Last Modification: 13/02/2010
'***************************************************
On Error GoTo ErrHandler

    If Cantidad > 0 Then
        Dim MiObj As Obj
        Dim RemainingGold As Long
    
        Cantidad = Cantidad * ConstantesBalance.ModGoldMultiplier
        
        RemainingGold = Cantidad
        
        While (RemainingGold > 0)
            
            ' Tira pilon de 10k
            If RemainingGold > MAX_INVENTORY_OBJS Then
                MiObj.Amount = MAX_INVENTORY_OBJS
                RemainingGold = RemainingGold - MAX_INVENTORY_OBJS
                
            ' Tira lo que quede
            Else
                MiObj.Amount = RemainingGold
                RemainingGold = 0
            End If

            MiObj.ObjIndex = ConstantesItems.Oro
            
            Call TirarItemAlPiso(Pos, MiObj)
        Wend
    End If

    Exit Sub

ErrHandler:
    Call LogError("Error en TirarOro. Error " & Err.Number & " : " & Err.Description)
End Sub

