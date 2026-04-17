VERSION 5.00
Begin VB.Form frmComerciar 
   BorderStyle     =   0  'None
   ClientHeight    =   7950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6915
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   530
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   461
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ARGENTUM.AOPictureBox picInvNpc 
      Height          =   3345
      Left            =   570
      TabIndex        =   3
      Top             =   1710
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   5900
   End
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3240
      TabIndex        =   2
      Text            =   "1"
      Top             =   3390
      Width           =   555
   End
   Begin ARGENTUM.AOPictureBox picInvUser 
      Height          =   3345
      Left            =   4050
      TabIndex        =   4
      Top             =   1710
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   5900
   End
   Begin VB.Image imgCerrar 
      Height          =   525
      Left            =   2850
      MouseIcon       =   "frmComerciar.frx":0000
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   7140
      Width           =   1230
   End
   Begin VB.Image imgVender 
      Height          =   270
      Left            =   3210
      MouseIcon       =   "frmComerciar.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "frmComerciar.frx":045C
      Tag             =   "1"
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image imgComprar 
      Height          =   270
      Left            =   3555
      MouseIcon       =   "frmComerciar.frx":136F
      MousePointer    =   99  'Custom
      Picture         =   "frmComerciar.frx":14C1
      Tag             =   "1"
      Top             =   3000
      Width           =   270
   End
   Begin VB.Label lblItemDescription 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ItemDescription"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   1995
      TabIndex        =   1
      Top             =   5880
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblItemName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ItemName"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Left            =   1920
      TabIndex        =   0
      Top             =   5595
      Visible         =   0   'False
      Width           =   3060
   End
End
Attribute VB_Name = "frmComerciar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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

Private clsFormulario As clsFormMovementManager

Public LastIndex1 As Integer
Public LastIndex2 As Integer
Public LasActionBuy As Boolean
Private ClickNpcInv As Boolean

Private cBotonVender As clsGraphicalButton
Private cBotonComprar As clsGraphicalButton
Private cBotonCruz As clsGraphicalButton

' WithEvents not allowed in arrays.
Private WithEvents dragNPCInventory As clsGraphicalInventory
Attribute dragNPCInventory.VB_VarHelpID = -1
Private WithEvents dragUserInventory As clsGraphicalInventory
Attribute dragUserInventory.VB_VarHelpID = -1

Public LastButtonPressed As clsGraphicalButton

Private Sub dragNPCInventory_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer, _
                                      ByVal moveType As eMoveType)
On Error GoTo ErrHandler

    If moveType = eMoveType.InventoryToTarget Then
        Call imgVender_Click
    End If
    
    Me.Refresh
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub dragNPCInventory_dragDone de frmComerciar.frm")
End Sub

Private Sub dragUserInventory_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer, _
                                       ByVal moveType As eMoveType)
On Error GoTo ErrHandler

    Select Case moveType
        '  Case eMoveType.Inventory
        '    Call Protocol.WriteMoveItem(originalSlot, newSlot, eMoveType.Inventory)
        
        Case eMoveType.InventoryToTarget
            Call imgComprar_Click
    End Select
    
    Me.Refresh
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub dragUserInventory_dragDone de frmComerciar.frm")
End Sub

Private Sub dragUserInventory_updateSlotInfo(ByVal nSlot As Integer)
    ' Current slot.
On Error GoTo ErrHandler
  
    Call picInvUser_Click
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub dragUserInventory_updateSlotInfo de frmComerciar.frm")
End Sub

Private Sub cantidad_Change()
On Error GoTo ErrHandler
    If Val(cantidad.Text) < 1 Then
        cantidad.Text = 1
    End If
    
    If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
        cantidad.Text = MAX_INVENTORY_OBJS
    End If
    
    If ClickNpcInv Then
        If InvComNpc.SelectedItem <> 0 Then
            'El precio, cuando nos venden algo, lo tenemos que redondear para arriba.
            'Label1(1).Caption = "Precio: " & CalculateSellPrice(NPCInventory(InvComNpc.SelectedItem).Valor, Val(cantidad.Text))  'No mostramos numeros reales
            Call InvComNpc.SelectItem(InvComNpc.SelectedItem)
            Call picInvNpc_Click
        End If
    Else
        If InvComUsu.SelectedItem <> 0 Then
            'Label1(1).Caption = "Precio: " & CalculateBuyPrice(Inventario.Valor(InvComUsu.SelectedItem), Val(cantidad.Text))  'No mostramos numeros reales
            Call InvComUsu.SelectItem(InvComUsu.SelectedItem)
            Call picInvUser_Click
        End If
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub cantidad_Change de frmComerciar.frm")
End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
On Error GoTo ErrHandler
  
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub cantidad_KeyPress de frmComerciar.frm")
End Sub

Private Sub Form_Load()

    ' Handles Form movement (drag and drop).
On Error GoTo ErrHandler
  
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me, , False
    
    Set dragUserInventory = InvComUsu
    Set dragNPCInventory = InvComNpc
    
    Set dragUserInventory = InvComNpc
    Set dragNPCInventory = InvComUsu
    
    'Cargamos la interfase
    Me.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "VentanaComerciar.jpg")
    
    Call LoadButtons
    
    Call modCustomCursors.SetFormCursorDefault(Me)

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmComerciar.frm")
End Sub

Private Sub LoadButtons()
On Error GoTo ErrHandler
  
    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI

    Set cBotonVender = New clsGraphicalButton
    Set cBotonComprar = New clsGraphicalButton
    Set cBotonCruz = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonVender.Initialize(imgVender, GrhPath & "BotonFlechaIzquierda_2.jpg", _
                                    GrhPath & "BotonFlechaIzquierda_2.jpg", _
                                    GrhPath & "BotonFlechaIzquierda_2.jpg", Me)

    Call cBotonComprar.Initialize(imgComprar, GrhPath & "BotonFlechaDerecha_2.jpg", _
                                    GrhPath & "BotonFlechaDerecha_2.jpg", _
                                    GrhPath & "BotonFlechaDerecha_2.jpg", Me)

    Call cBotonCruz.Initialize(imgCerrar, GrhPath & "BotonCerrar.jpg", _
                                    GrhPath & "BotonCerrar.jpg", _
                                    GrhPath & "BotonCerrar.jpg", Me)

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadButtons de frmComerciar.frm")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

''
' Calculates the selling price of an item (The price that a merchant will sell you the item)
'
' @param objValue Specifies value of the item.
' @param objAmount Specifies amount of items that you want to buy
' @return   The price of the item.

Private Function CalculateSellPrice(ByRef objValue As Single, ByVal objAmount As Long) As Double
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 19/08/2008
'Last modify by: Franco Zeoli (Noich)
'*************************************************
    On Error GoTo Error
    'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
    CalculateSellPrice = CCur(objValue * 1000000) / 1000000 * objAmount
    
    Exit Function
Error:
    MsgBox Err.Description, vbExclamation, "Error: " & Err.Number
End Function
''
' Calculates the buying price of an item (The price that a merchant will buy you the item)
'
' @param objValue Specifies value of the item.
' @param objAmount Specifies amount of items that you want to buy
' @return   The price of the item.
Private Function CalculateBuyPrice(ByRef objValue As Single, ByVal objAmount As Long) As Double
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 19/08/2008
'Last modify by: Franco Zeoli (Noich)
'*************************************************
    On Error GoTo Error
    'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
    CalculateBuyPrice = Fix(CCur(objValue * 1000000) / 1000000 * objAmount)
    
    Exit Function
Error:
    MsgBox Err.Description, vbExclamation, "Error: " & Err.Number
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set dragNPCInventory = Nothing
On Error GoTo ErrHandler
  
    Set dragUserInventory = Nothing
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Unload de frmComerciar.frm")
End Sub

Private Sub imgComprar_Click()
    ' Debe tener seleccionado un item para comprarlo.
On Error GoTo ErrHandler
  
    If InvComNpc.SelectedItem = 0 Then Exit Sub
    
    If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
    
    Call Engine_Audio.PlayInterface(SND_CLICK)
    
    If InvComNpc.ObjIndex(InvComNpc.SelectedItem) <= 0 Then Exit Sub
    
    LasActionBuy = True
    If UserGLD >= CalculateSellPrice(NPCInventory(InvComNpc.SelectedItem).Valor, Val(cantidad.Text)) Then
        Call WriteCommerceBuy(InvComNpc.SelectedItem, Val(cantidad.Text))
    Else
        Call AddtoRichTextBox(frmMain.RecTxt(0), "No tienes suficiente oro.", 2, 51, 223, 1, 1, , eMessageType.Info)
        Exit Sub
    End If
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgComprar_Click de frmComerciar.frm")
End Sub

Private Sub CerrarVentana()
    Call dragNPCInventory.Release
    Call dragUserInventory.Release
    
    Set dragNPCInventory = Nothing
    Set dragUserInventory = Nothing
    
    Call WriteCommerceEnd
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CerrarVentana
End Sub

Private Sub imgCerrar_Click()
On Error GoTo ErrHandler
  
    Call CerrarVentana
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgCross_Click de frmComerciar.frm")
End Sub

Private Sub imgVender_Click()
    ' Debe tener seleccionado un item para comprarlo.
On Error GoTo ErrHandler
  
    If InvComUsu.SelectedItem = 0 Then Exit Sub
    
    If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
    
    If InvComUsu.ObjIndex(InvComUsu.SelectedItem) <= 0 Then Exit Sub
    
    Call Engine_Audio.PlayInterface(SND_CLICK)
    
    LasActionBuy = False

    Call WriteCommerceSell(InvComUsu.SelectedItem, Val(cantidad.Text))
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgVender_Click de frmComerciar.frm")
End Sub


Private Sub picInvNpc_Click()
    If dragNPCInventory.Dragging Then Exit Sub
On Error GoTo ErrHandler
  
    Dim ItemSlot As Byte
    Dim ItemMinLevel As Byte
    Dim ShowItemLevel As Boolean
    Dim ItemDescription As String
    Dim ItemPrice As String
    
    ItemSlot = InvComNpc.SelectedItem
    
    ClickNpcInv = True
    InvComUsu.DeselectItem
    
    If ItemSlot <= 0 Then
        lblItemName.Caption = ""
        lblItemDescription.Visible = False
        InvComUsu.DeselectItem
        Exit Sub
    End If
    
    lblItemName.Caption = NPCInventory(ItemSlot).Name
    lblItemName.Visible = True
            
    ItemMinLevel = GameMetadata.Objs(NPCInventory(ItemSlot).ObjIndex).MinimumLevel
    ItemPrice = "Precio: " & CalculateSellPrice(NPCInventory(ItemSlot).Valor, Val(cantidad.Text)) 'No mostramos numeros reales
    
    Select Case NPCInventory(ItemSlot).OBJType
        Case eObjType.otWeapon, eObjType.otTool
            ItemDescription = "Golpe Mínimo: " & NPCInventory(ItemSlot).MinHit & vbCrLf & "Golpe Máximo: " & NPCInventory(ItemSlot).MaxHit
            ShowItemLevel = True
        Case eObjType.otArmadura, eObjType.otCasco, eObjType.otEscudo
            ItemDescription = "Defensa Mínima: " & NPCInventory(ItemSlot).MinDef & vbCrLf & "Defensa Máxima: " & NPCInventory(ItemSlot).MaxDef
            ShowItemLevel = True
        Case Else
            lblItemDescription.Visible = False
    End Select
    
    If ShowItemLevel And ItemMinLevel > 0 Then
        ItemDescription = ItemDescription & vbCrLf & "Nivel Mínimo: " & ItemMinLevel
    End If
    
    ItemDescription = ItemPrice & vbCrLf & ItemDescription
    lblItemDescription.Caption = ItemDescription
    lblItemDescription.Visible = True
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub picInvNpc_Click de frmComerciar.frm")
End Sub

Private Sub picInvNpc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Public Function getCount() As Long
On Error GoTo ErrHandler
  
    getCount = cantidad.Text
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function getCount de frmComerciar.frm")
End Function
Private Sub picInvUser_Click()
On Error GoTo ErrHandler
      
    Dim ItemSlot As Byte
    Dim ItemMinLevel As Byte
    Dim ShowItemLevel As Boolean
    Dim ItemDescription As String
    Dim ItemPrice As String
    
    If dragUserInventory.Dragging Then Exit Sub
    
    ItemSlot = InvComUsu.SelectedItem
    
    InvComNpc.DeselectItem
    ClickNpcInv = False
        
    If ItemSlot <= 0 Then
        lblItemName.Caption = ""
        lblItemDescription.Visible = False

        Exit Sub
    End If
    
    lblItemName.Caption = Inventario.ItemName(ItemSlot)
    lblItemName.Visible = True
    
    ItemMinLevel = GameMetadata.Objs(Inventario.ObjIndex(ItemSlot)).MinimumLevel
    ItemPrice = "Precio: " & CalculateSellPrice(Inventario.Valor(ItemSlot), Val(cantidad.Text)) 'No mostramos numeros reales
    
    Select Case Inventario.OBJType(ItemSlot)
        Case eObjType.otWeapon, eObjType.otTool
            ItemDescription = "Golpe Mínimo: " & Inventario.MinHit(ItemSlot) & vbCrLf & "Golpe Máximo: " & Inventario.MaxHit(ItemSlot)
            ShowItemLevel = True
        Case eObjType.otArmadura, eObjType.otCasco, eObjType.otEscudo
            ItemDescription = "Defensa Mínima: " & Inventario.MinDef(ItemSlot) & vbCrLf & "Defensa Máxima: " & Inventario.MaxDef(ItemSlot)
            ShowItemLevel = True
        Case Else
            lblItemDescription.Visible = False
    End Select
    
    If ShowItemLevel And ItemMinLevel > 0 Then
        ItemDescription = ItemDescription & vbCrLf & "Nivel Mínimo: " & ItemMinLevel
    End If
    
    ItemDescription = ItemPrice & vbCrLf & ItemDescription
    lblItemDescription.Caption = ItemDescription
    lblItemDescription.Visible = True
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub picInvUser_Click de frmComerciar.frm")
End Sub

Private Sub picInvUser_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub
