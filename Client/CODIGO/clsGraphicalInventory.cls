VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Graphical inventory for ORE 0.4.1
'Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
'juansotuyo@hotmail.com
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
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

'Private invFont As tFuente
Private InventoryOffset As Long             'Number of lines we scrolled down from topmost
Private InvSelectedItem As Long             'Currently selected item

Private ownTilePixelWidth As Integer
Private ownTilePixelHeight As Integer
Private lastTileRow As Integer

Private ShowValue As Boolean
Private ValueFormated As Boolean

Private ItemSeparatorSize As Byte

Private ImgContainer As Boolean

Public slotDragged As Integer

Private dragAndDropOn As Boolean

Private inventory() As inventory    'User's inventory

Public WithEvents InventoryWindow As AOPictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1
Public Event dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer, ByVal moveType As eMoveType)

Public dropInventory As clsGraphicalInventory
Private inventoryMoveType As eMoveType

Private Device As Long

Private Const INV_SELECTOR_RED_GRH As Integer = 8553
Private Const INV_SELECTOR_GREEN_GRH As Integer = 8554
Private Const INV_SELECTOR_BLUE_GRH As Integer = 8555

Private Invalidated As Boolean

''' Event Handler
Private InternalEventHandler As clsButtonEventHandler
Private InternalIndex As Integer

Public Property Get Name() As String
    Name = InventoryWindow.Name
End Property

Public Property Get Index() As Integer
    Index = InternalIndex
End Property

Public Property Let Index(NewIndex As Integer)
    InternalIndex = NewIndex
End Property

Public Property Get EventHandler() As clsButtonEventHandler
    Set EventHandler = InternalEventHandler
End Property

Public Property Set EventHandler(eh As clsButtonEventHandler)
    Set InternalEventHandler = eh
End Property
'''

Private Sub Class_Initialize()
    Invalidated = False
End Sub

Private Sub Class_Terminate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
On Error GoTo ErrHandler
    
    Call Aurora_Graphics.DeletePass(Device)
    
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Class_Terminate de clsGrapchicalInventory.cls")
End Sub


Public Sub ScrollInventory(ByVal Up As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 14/01/2010
'Scrolls the graphical inventory up / down
'14/01/2010: ZaMa - Now redraws inventory only if needed
'***************************************************
On Error GoTo ErrHandler
  

    'Only move if there is something to show up /down
    If CanScroll(Up) Then
        InventoryOffset = InventoryOffset + IIf(Up, 1, -1)
        
        'Redraw the inventory if needed
        Call Invalidate
    End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ScrollInventory de clsGrapchicalInventory.cls")
End Sub

Public Function CanScroll(ByVal Up As Boolean)
'***************************************************
'Author: ZaMa
'Last Modify Date: 14/01/2010
'Determines wether inventory can be scrolled up / down
'***************************************************
On Error GoTo ErrHandler
  
    If Up Then
        CanScroll = (InventoryOffset + (InventoryWindow.ScaleHeight \ ownTilePixelHeight)) * _
                    (InventoryWindow.ScaleWidth \ ownTilePixelWidth) < UBound(inventory)
    Else
        CanScroll = InventoryOffset > 0
    End If
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CanScroll de clsGrapchicalInventory.cls")
End Function

Public Sub SelectGold()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Sets the gold as the currently selected item
'***************************************************
On Error GoTo ErrHandler
  
    Dim prevSelectedItem As Long
    
    'Store preivously selected item
    prevSelectedItem = InvSelectedItem
    
    'Select the gold
    InvSelectedItem = FLAGORO
    
    'Redraw old item to deselect it
    If prevSelectedItem <> FLAGORO Then
        Call Invalidate
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SelectGold de clsGrapchicalInventory.cls")
End Sub

Public Sub SelectItem(ByVal Slot As Integer)
    
    If Slot = InvSelectedItem Then Exit Sub
    
    If Slot < 1 Or Slot > UBound(inventory) Then
        InvSelectedItem = 0
    End If
        
    InvSelectedItem = Slot
    
    'Invalidate
    Call Invalidate

End Sub

Public Sub DeselectItem()
'***************************************************
'Author: ZaMa
'Last Modify Date: 30/11/2009
'Deselect the currently selected item
'***************************************************
On Error GoTo ErrHandler
  
    If InvSelectedItem = 0 Then Exit Sub
    
    Dim ClearSlot As Byte
    
    ClearSlot = InvSelectedItem
    
    'Select nothing
    InvSelectedItem = 0
    
    'Redraw old item to deselect it
    Call Invalidate

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DeselectItem de clsGrapchicalInventory.cls")
End Sub

Public Sub ChangeSlotItemPrice(ByVal Slot As Byte, ByVal Price As Long)
On Error GoTo ErrHandler
  
    inventory(Slot).Valor = Price
    
    'Update Price
    Call Invalidate
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ChangeSlotItemPrice de clsGrapchicalInventory.cls")
End Sub

Public Sub ChangeSlotItemAmount(ByVal Slot As Byte, ByVal Amount As Long)
'***************************************************
'Author: ZaMa
'Created Date: 30/11/2009
'Last Modify Date: 09/12/2009 (Budi)
'Change the amount of the required item, and update the slot.
'***************************************************
On Error GoTo ErrHandler
  
    
    inventory(Slot).Amount = Amount
    
    'Update Amount
    Call Invalidate
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ChangeSlotItemAmount de clsGrapchicalInventory.cls")
End Sub

Public Function Dragging() As Boolean
'***************************************************
'Author: D'Artagnan
'Last Modify Date: 28/10/2014
'Return True if an item is being dragged. False otherwise.
'***************************************************
On Error GoTo ErrHandler
  
    Dragging = CBool(slotDragged)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function Dragging de clsGrapchicalInventory.cls")
End Function

Public Property Get SelectedItem() As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item index
'***************************************************
    SelectedItem = InvSelectedItem
End Property

Public Property Get MaxHit(ByVal Slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the max hit of the selected item
'***************************************************
    MaxHit = inventory(Slot).MaxHit
End Property

Public Property Get MinHit(ByVal Slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the min hit of the selected item
'***************************************************
    MinHit = inventory(Slot).MinHit
End Property

Public Property Get MaxDef(ByVal Slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the defense of the selected item
'***************************************************
    MaxDef = inventory(Slot).MaxDef
End Property

Public Property Get MinDef(ByVal Slot As Byte) As Integer
'***************************************************
'Author: Budi
'Last Modify Date: 02/1209
'Retrieves the defense of the selected item
'***************************************************
    MinDef = inventory(Slot).MinDef
End Property

Public Property Get GrhIndex(ByVal Slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the grh index of the selected item
'***************************************************
    GrhIndex = inventory(Slot).GrhIndex
End Property

Public Property Get Valor(ByVal Slot As Byte) As Single
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the value of the selected item
'***************************************************
    Valor = inventory(Slot).Valor
End Property

Public Property Get Amount(ByVal Slot As Byte) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's amount
'***************************************************
    If Slot = FLAGORO Then
        Amount = UserGLD
    ElseIf Slot >= LBound(inventory) And Slot <= UBound(inventory) Then
        Amount = inventory(Slot).Amount
    End If
End Property

Public Property Get ObjIndex(ByVal Slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object index
'***************************************************
    ObjIndex = inventory(Slot).ObjIndex
End Property

Public Property Get OBJType(ByVal Slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object type
'***************************************************
    OBJType = inventory(Slot).OBJType
End Property

Public Property Get ItemName(ByVal Slot As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's name
'***************************************************
    ItemName = inventory(Slot).Name
End Property

Public Property Get Equipped(ByVal Slot As Byte) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves True if the item at the given pos is eqiupped
'***************************************************
    Equipped = inventory(Slot).Equipped
End Property

Public Property Get CanUse(ByVal indexSlot As Byte) As Boolean
'Autor: G Toyz
'Fecha: 5/17 (mm/aa)
    CanUse = inventory(indexSlot).CanUse
End Property

Public Property Get MaxObjs() As Byte
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 09/16/09
'Retrieves the capacity of the Inventory
'***************************************************
    MaxObjs = UBound(inventory)
End Property

Public Property Get UsableHeight() As Long
'***************************************************
'Author: D'Artagnan
'Last Modify Date: 24/08/14
'The inventory usable height (where an item might be), in pixels.
'***************************************************
    UsableHeight = MaxObjs / 5 * ownTilePixelHeight
End Property

Public Property Get Height() As Long
'***************************************************
'Author: D'Artagnan
'Last Modify Date: 24/08/14
'The inventory height, in pixels.
'***************************************************
    Height = InventoryWindow.ScaleHeight
End Property

Public Property Get Width() As Long
'***************************************************
'Author: D'Artagnan
'Last Modify Date: 24/08/14
'The inventory width, in pixels.
'***************************************************
    Width = InventoryWindow.ScaleWidth
End Property

Public Property Get Left() As Long
'***************************************************
'Author: D'Artagnan
'Last Modify Date: 24/08/14
'The inventory X position, in pixels.
'***************************************************
    Left = InventoryWindow.Left
End Property

Public Property Get Top() As Long
'***************************************************
'Author: D'Artagnan
'Last Modify Date: 24/08/14
'The inventory Y position, in pixels.
'***************************************************
    Top = InventoryWindow.Top
End Property

Public Sub SetMaxObjs(ByVal MaxObjs As Byte)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 09/16/09
'Set the capacity of the Inventary
'***************************************************
On Error GoTo ErrHandler
  
    If UBound(inventory) = MaxObjs Then Exit Sub
    
    ReDim Preserve inventory(1 To MaxObjs) As inventory
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SetMaxObjs de clsGrapchicalInventory.cls")
End Sub

Public Sub ToggleShowValue(ByVal bShow As Boolean)
'***************************************************
'Author: ZaMa
'Last Modify Date: 11/12
'Indicates if the amount of the items in the inventory is shown
'***************************************************
On Error GoTo ErrHandler
  
    ShowValue = bShow
    
    Call Invalidate
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ToggleShowValue de clsGrapchicalInventory.cls")
End Sub

Public Sub ClearAllSlots()
'***************************************************
'Author: ZaMa
'Last Modify Date: 15/03/2011
'Clear All Slots.
'***************************************************
On Error GoTo ErrHandler
  

    Dim Slot As Long
    For Slot = 1 To Me.MaxObjs
        Call SetItem(Slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString, 0, True)
    Next Slot

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ClearAllSlots de clsGrapchicalInventory.cls")
End Sub

Public Sub SetItem(ByVal Slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Long, ByVal eEquipped As Byte, _
    ByVal eGrhIndex As Integer, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
    ByVal MaxDef As Integer, ByVal MinDef As Integer, ByVal eValor As Single, ByVal eName As String, _
    ByVal MaxAmount As Long, Optional ByVal CanUse As Boolean = True)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/04/06
'Sets all data for a given inventory slot
'***************************************************
On Error GoTo ErrHandler
  
    If Slot < 1 Or Slot > UBound(inventory) Then Exit Sub
    
    With inventory(Slot)
        .Amount = eAmount
        .MaxDef = MaxDef
        .MinDef = MinDef
        .Equipped = eEquipped
        .GrhIndex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .Name = eName
        .ObjIndex = eOBJIndex
        .OBJType = eObjType
        .Valor = eValor
        .MaxAmount = MaxAmount
        .CanUse = CanUse
    End With
    
    'If InventoryWindow.Visible = False Then InventoryWindow.Visible = True
    
    Call Invalidate
        
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SetItem de clsGrapchicalInventory.cls")
End Sub

Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Selects the item clicked if it's valid and return's it's index
'***************************************************
On Error GoTo ErrHandler
  
    Dim TempItem As Long
    Dim temp_x As Long
    Dim temp_y As Long
    
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
    
    'Make sure it's within limits
    If TempItem > 0 And TempItem <= UBound(inventory) Then
        'Make sure slot isn't empty
        If inventory(TempItem).GrhIndex Then
            ClickItem = TempItem
        Else
            ClickItem = 0
        End If
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ClickItem de clsGrapchicalInventory.cls")
End Function

Private Sub InventoryWindow_Paint()
    Call DrawInventory
End Sub

Public Sub DrawInventory()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Renders the inventory to the given PictureBox
'***************************************************
On Error GoTo ErrHandler
  
    Dim I As Long
  
    If InventoryWindow Is Nothing Then
        Exit Sub
    End If

    'Clear the inventory window
    Call UIBegin(Device, InventoryWindow.ScaleWidth, InventoryWindow.ScaleHeight, &H0)
              
    For I = 0 To UBound(inventory) - 1
        With inventory(I + 1)
            If .GrhIndex Then

                Dim DestX As Long, DestY As Long, PositionInRow As Byte, PositionInColumn As Byte
                PositionInRow = (I Mod lastTileRow)
                PositionInColumn = (I \ lastTileRow)
                
                DestX = PositionInRow * ownTilePixelWidth
                DestY = PositionInColumn * ownTilePixelHeight
                
                If ItemSeparatorSize > 0 Then
                    DestX = DestX + PositionInRow + ItemSeparatorSize
                    DestY = DestY + PositionInColumn + ItemSeparatorSize
                End If
                                
                'Render the item grh and the amount
                If Not .CanUse Then
                    Call DrawGrhIndex(.GrhIndex, DestX, DestY, GetDepth(1, DestX, DestY, 1), 0, &HFF0000FF)
                Else
                    Call DrawGrhIndex(.GrhIndex, DestX, DestY, GetDepth(1, DestX, DestY, 1), 0, &HFFFFFFFF)
                End If
               
                If .Amount > 0 Then Call DrawText(DestX + 1, DestY + 1, GetDepth(1, DestX, DestY, 2), CStr(.Amount), FuentesJuego.Inventarios.color, eRendererAlignmentLeftTop, FuentesJuego.Inventarios, True)
               
                ' If equipped we render "E"
                If .Equipped Then
                    Call DrawText(DestX + ownTilePixelWidth - 10, DestY + ownTilePixelHeight - 12, GetDepth(1, DestX, DestY, 3), "E", RGBA(255, 255, 0, 255), eRendererAlignmentLeftTop, FuentesJuego.Inventarios, True)
                End If
                
                ' Render a box around the selected item
                If InvSelectedItem = (I + 1) And ImgContainer Then
                    Call DrawGrhIndex(INV_SELECTOR_GREEN_GRH, DestX, DestY, GetDepth(1, DestX, DestY, 4), 0)
                End If
            End If
        End With
    Next I

    Call UIEnd
        
    Invalidated = False
    
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DrawInventory de clsGrapchicalInventory.cls")
End Sub

Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse move event of the inventory picture box
'Displays a ToolTip of the item under the mouse
'***************************************************
    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim TempItem As Integer

    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
        
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
    
    If TempItem <= UBound(inventory) And TempItem > 0 Then
        InventoryWindow.ToolTipText = IIf(LenB(inventory(TempItem).Name) > 0, inventory(TempItem).Name, "Nada")
    End If

End Sub

Private Sub InventoryWindow_DragBegin(ByVal Source As AOPictureBox, ByVal Shift As Boolean, ByVal X As Single, ByVal Y As Single, ByRef Cancelled As Boolean)
    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim TempItem As Integer

    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
        
    TempItem = temp_x + (temp_y + InventoryOffset) * (Me.Width \ ownTilePixelWidth) + 1
            
    If DragAllowed(TempItem) Then
        slotDragged = TempItem
    Else
        Cancelled = True
    End If
    
End Sub
Private Sub InventoryWindow_DragEnd(ByVal Source As AOPictureBox, ByVal Shift As Boolean, ByVal X As Single, ByVal Y As Single)
    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim TempItem As Integer
On Error GoTo ErrHandler
  
    Dim nDropSlot As Integer

    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
        
    nDropSlot = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
            
    If InvSelectedItem <= MaxObjs Then
        If Source.hwnd = InventoryWindow.hwnd Then
                If nDropSlot <> slotDragged Then
                    InvSelectedItem = nDropSlot
                        
                    Call Me.RaiseDragDoneEvent(slotDragged, nDropSlot, eMoveType.inventory)
                End If
        End If
        
        If Not dropInventory Is Nothing Then
            If Source.hwnd = dropInventory.InventoryWindow.hwnd Then
                Call dropInventory.RaiseDragDoneEvent(slotDragged, nDropSlot, eMoveType.InventoryToTarget)
            End If
        End If
        
        slotDragged = 0
    End If

    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub InventoryWindow_DragDrop de clsGrapchicalInventory.cls")
End Sub



Private Function DragAllowed(ByVal nSlot As Integer) As Boolean
'***************************************************
'Author: D'Artagnan (based on Budi's code)
'Last Modification: 18/08/2014
'Return True when D&D is allowed at the specified slot. False otherwise.
'***************************************************
On Error GoTo ErrHandler
  
    Dim upperLeftSlot As Integer
    
    upperLeftSlot = InventoryOffset * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
    
    If nSlot > (UBound(inventory) - 1) Then
        DragAllowed = False
    Else
        DragAllowed = IIf(nSlot >= upperLeftSlot And nSlot <= MaxObjs, inventory(nSlot).GrhIndex, False)
    End If
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function DragAllowed de clsGrapchicalInventory.cls")
End Function


Friend Sub Initialize(ByRef InvPic As AOPictureBox, ByVal MaxObjs As Byte, _
                          Optional ByVal FontSize As Integer = 7, Optional ByVal TileWidth As Integer = 32, _
                          Optional ByVal TileHeight As Integer = 32, Optional ByVal startX As Integer = 0, _
                          Optional ByVal startY As Integer = 0, Optional ByVal bImgContainer As Boolean = True, _
                          Optional ByVal bShowText As Boolean = True, Optional FormatValue As Boolean = False, _
                          Optional ByVal dragAndDrop As Boolean = False, _
                          Optional ByVal moveType As eMoveType, _
                          Optional ByRef dropInventoryL As clsGraphicalInventory, _
                          Optional ByVal ItemSeparatorSizeInPixels As Byte = 0)
                          
On Error GoTo ErrHandler

'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 24/08/14
'Sets the reference to the picture box on which to render the inventory
'03/12/09: I added the optionals args FontSize, TileWidth, TileHeight, startX and startY. (Budi)
'24/08/14: Optional arguments: moveType and dropInventoryL. (D'Artagnan)
'***************************************************
        
    Set InternalEventHandler = Nothing
    Set InventoryWindow = InvPic
    ValueFormated = FormatValue

    ReDim inventory(1 To MaxObjs) As inventory
      
    ownTilePixelHeight = TileHeight
    ownTilePixelWidth = TileWidth
    lastTileRow = InvPic.ScaleWidth \ TileWidth
    
    ImgContainer = bImgContainer
    ShowValue = bShowText
    
    ItemSeparatorSize = ItemSeparatorSizeInPixels
   
    dragAndDropOn = dragAndDrop
    If dragAndDrop Then
        inventoryMoveType = moveType
        Set dropInventory = dropInventoryL
    End If

    Device = Aurora_Graphics.CreatePassFromDisplay(InvPic.hwnd, InvPic.ScaleWidth, InvPic.ScaleHeight)
    
    If ImgContainer Then _
        InvSelectedItem = ClickItem(1, 1)   'If there is anything there we select the top left item

    Call Invalidate
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Initialize de clsGrapchicalInventory.cls")
End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim prevSelItem As Long

    prevSelItem = InvSelectedItem
    
    'Get the currently clickced item
    InvSelectedItem = ClickItem(CInt(X), CInt(Y))
    
    '''
    ' Raise the click event on the attached event handler.
    ' This is normally used when setting event handlers from an array of controls.
    If Not (InternalEventHandler Is Nothing) Then
        Call InternalEventHandler.OnGraphicalInventoryClick(Me)
    End If
    '''
        
    'Update needed inventory slots
    If prevSelItem <> InvSelectedItem Then
        If prevSelItem <> 0 And prevSelItem <> FLAGORO Or InvSelectedItem Then
            Call Invalidate
        End If
    End If

End Sub

Public Sub RaiseDragDoneEvent(ByVal nOriginalSlot As Integer, ByVal nNewSlot As Integer, ByVal moveType As eMoveType)
'***************************************************
'Author: D'Artagnan
'Last Modify Date: 24/08/14
'
'***************************************************
On Error GoTo ErrHandler
  
    RaiseEvent dragDone(nOriginalSlot, nNewSlot, moveType)

    slotDragged = 0

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RaiseDragDoneEvent de clsGrapchicalInventory.cls")
End Sub

Public Sub Release()
'***************************************************
'Author: D'Artagnan
'Last Modify Date: 24/08/14
'This should be called before deleting the reference
'of a class instance (setting its value to Nothing).
'***************************************************
On Error GoTo ErrHandler
  
    Set dropInventory = Nothing
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Release de clsGrapchicalInventory.cls")
End Sub

Private Sub Invalidate()
    If Not Invalidated Then
        Call Mod_General.Invalidate(InventoryWindow.hwnd)
        Invalidated = True
    End If
End Sub




