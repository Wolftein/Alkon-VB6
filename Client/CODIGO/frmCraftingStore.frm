VERSION 5.00
Begin VB.Form frmCraftingStore 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Tienda de Construcción"
   ClientHeight    =   6510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "frmCraftingStore.frx":0000
   ScaleHeight     =   6510
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ARGENTUM.AOPictureBox picItemsToCraft 
      Height          =   875
      Left            =   5510
      TabIndex        =   7
      Top             =   4950
      Width           =   3875
      _ExtentX        =   6826
      _ExtentY        =   1535
   End
   Begin ARGENTUM.AOPictureBox picRequiredMaterials 
      Height          =   875
      Left            =   600
      TabIndex        =   6
      Top             =   5260
      Width           =   3875
      _ExtentX        =   6826
      _ExtentY        =   1535
   End
   Begin ARGENTUM.AOPictureBox picCraftableItems 
      Height          =   3050
      Left            =   600
      TabIndex        =   5
      Top             =   1870
      Width           =   3890
      _ExtentX        =   6853
      _ExtentY        =   5371
   End
   Begin VB.TextBox txtCraftingStoreOpenLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Height          =   735
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmCraftingStore.frx":10C7D
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtMaterialsPrice 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   230
      Left            =   8140
      TabIndex        =   3
      Text            =   "0"
      Top             =   2870
      Width           =   1140
   End
   Begin VB.TextBox txtConstructionPrice 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   230
      Left            =   6110
      TabIndex        =   2
      Text            =   "0"
      Top             =   2870
      Width           =   1140
   End
   Begin VB.Image imgHistoryButton 
      Height          =   525
      Left            =   7200
      Top             =   3600
      Width           =   735
   End
   Begin VB.Image imgClose 
      Height          =   330
      Left            =   9960
      Top             =   180
      Width           =   330
   End
   Begin VB.Image imgOpenStoreButton 
      Height          =   525
      Left            =   8280
      Top             =   3600
      Width           =   1230
   End
   Begin VB.Image imgCraftButton 
      Height          =   525
      Left            =   6960
      Top             =   3600
      Width           =   1230
   End
   Begin VB.Image imgAddRecipeToCraftButton 
      Height          =   525
      Left            =   5640
      Top             =   3600
      Width           =   1230
   End
   Begin VB.Label lblItemName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Objeto"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   5880
      TabIndex        =   1
      Top             =   1710
      Width           =   3255
   End
   Begin VB.Label lblWorkerName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Trabajador"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmCraftingStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tCraftableItem
    CraftingGroup As Byte
    RecipeIndex As Integer
    GrhIndex As Integer
    ObjIndex As Integer
End Type

Private TempCraftingRecipeGroups() As tCraftingRecipeGroup
Private TempCraftingRecipeGroupsQty As Integer

Private CraftableItemsSelected() As tCraftableItem

Private cButtonAddRecipe As clsGraphicalButton
Private cButtonCraft As clsGraphicalButton
Private cButtonOpenStore As clsGraphicalButton
Private cButtonHistory As clsGraphicalButton

Private cButtonClose As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private cForm As clsFormMovementManager

Public WithEvents InvCraftableItems As clsGraphicalInventory
Attribute InvCraftableItems.VB_VarHelpID = -1
Public WithEvents InvMaterialsRequired As clsGraphicalInventory
Attribute InvMaterialsRequired.VB_VarHelpID = -1

Public WithEvents InvItemsToCraft As clsGraphicalInventory
Attribute InvItemsToCraft.VB_VarHelpID = -1

Private WithEvents TabEventHandler As clsButtonEventHandler
Attribute TabEventHandler.VB_VarHelpID = -1

Private PicTabs() As VB.Image 'Picture boxes.
Private cPicTabs() As clsGraphicalButton
Private SelectedCraftingGroup As Byte

Const MAX_ITEMS As Integer = 100
Const MAX_ITEMS_MATERIALS As Integer = 20
Const MAX_ITEMS_TO_CRAFT As Byte = 20
Const MAX_PRICE As Long = 10000000


Private SelectedObjNumber As Integer
Private SelectedRecipeNumber As Integer
Private SelectedRecipeIndex As Integer

Dim CustomerMode As Boolean
Dim StoreInstanceId As String
Dim IsStoreOpen As Boolean


Public Sub ShowCraftingStoreData()

    lblWorkerName.Caption = CurrentOpenStore.OwnerName
    
End Sub

Private Sub CloseWindow()

    If CustomerMode = False And IsStoreOpen Then
        Call WriteWorkerStore_Close
        IsStoreOpen = False
    End If
    
    ViewingFormCantMove = False
    
    SelectedObjNumber = 0
    SelectedRecipeNumber = 0
    SelectedRecipeIndex = 0
    
    Call CleanMetadata
    
    Call InvItemsToCraft.ClearAllSlots
    
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Call CloseWindow
    
End Sub

Private Sub Form_Load()
    Call modCustomCursors.SetFormCursorDefault(Me)
    
    Call InitUI
    
    Call InitInventories
    
    Call LoadTabButtons
    
    If InvCraftableItems.MaxObjs > 0 Then Call InvCraftableItems.SelectItem(1)

    ViewingFormCantMove = True
End Sub

Public Sub InitUI()

    Dim ImgPath As String
    Set cForm = New clsFormMovementManager
    cForm.Initialize Me, , False
    
    txtCraftingStoreOpenLabel.Width = picCraftableItems.Width
    txtCraftingStoreOpenLabel.Height = picCraftableItems.Height
    txtCraftingStoreOpenLabel.Left = picCraftableItems.Left
    txtCraftingStoreOpenLabel.Top = picCraftableItems.Top
    txtCraftingStoreOpenLabel.Visible = False
    
    lblItemName.Caption = vbNullString
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Set cButtonAddRecipe = New clsGraphicalButton
    Set cButtonCraft = New clsGraphicalButton
    Set cButtonOpenStore = New clsGraphicalButton
    Set cButtonHistory = New clsGraphicalButton
    
    Set cButtonClose = New clsGraphicalButton
    
    ImgPath = DirInterfaces & SELECTED_UI
    Set Me.Picture = LoadPicture(ImgPath & "VentanaSelfWorker.jpg")
    
    Call cButtonAddRecipe.Initialize(imgAddRecipeToCraftButton, ImgPath & "BotonSelfWorker_Agregar.jpg", ImgPath & "BotonSelfWorker_Agregar.jpg", ImgPath & "BotonSelfWorker_Agregar.jpg", Me, ImgPath & "BotonSelfWorker_Agregar.jpg")
    Call cButtonCraft.Initialize(imgCraftButton, ImgPath & "BotonSelfWorker_Construir.jpg", ImgPath & "BotonSelfWorker_Construir.jpg", ImgPath & "BotonSelfWorker_Construir.jpg", Me, ImgPath & "BotonSelfWorker_Construir.jpg")
    Call cButtonOpenStore.Initialize(imgOpenStoreButton, ImgPath & "BotonSelfWorker_AbrirTienda.jpg", ImgPath & "BotonSelfWorker_AbrirTienda.jpg", ImgPath & "BotonSelfWorker_AbrirTienda.jpg", Me, ImgPath & "BotonSelfWorker_AbrirTienda.jpg")
    Call cButtonHistory.Initialize(imgHistoryButton, ImgPath & "BotonSelfWorker_History.jpg", ImgPath & "BotonSelfWorker_History.jpg", ImgPath & "BotonSelfWorker_History.jpg", Me, ImgPath & "BotonSelfWorker_History.jpg")
    
    Call cButtonClose.Initialize(imgClose, ImgPath & "BotonSelfWorker_Salir.jpg", ImgPath & "BotonSelfWorker_Salir.jpg", ImgPath & "BotonSelfWorker_Salir.jpg", Me, ImgPath & "BotonSelfWorker_Salir.jpg")
          
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Public Sub InitInventories()
On Error GoTo ErrHandler

    Set InvCraftableItems = New clsGraphicalInventory
    Set InvMaterialsRequired = New clsGraphicalInventory
    Set InvItemsToCraft = New clsGraphicalInventory
    
    Call InvCraftableItems.Initialize(frmCraftingStore.picCraftableItems, MAX_ITEMS, , , , 10, , , , , False)
    Call InvMaterialsRequired.Initialize(frmCraftingStore.picRequiredMaterials, MAX_ITEMS_MATERIALS, , , , 10, , , , , False)
    
    Call InvItemsToCraft.Initialize(frmCraftingStore.picItemsToCraft, MAX_ITEMS, , , , 10, , , , , False)
            
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub InitInventories de frmCraftingStore.frm")
End Sub

Public Sub SetStoreMode(ByVal IsCustomer As Boolean, ByRef StoreOwnerName As String, Optional ByRef InstanceId As String = vbNullString)
    CustomerMode = IsCustomer
    lblWorkerName.Caption = StoreOwnerName
    
    StoreInstanceId = InstanceId
    Call InitControls
End Sub

Public Sub SetStoreStatus(ByVal OpenState As Boolean, ByVal ChangeFormControlProperties As Boolean)
    IsStoreOpen = OpenState
  
    If ChangeFormControlProperties Then
        Call frmCraftingStore.SetFormControlsState(False, CustomerMode)
    End If
    
    Dim I As Integer
    For I = 1 To PlayerData.CraftingRecipeGroupsQty
        PicTabs(I).Visible = Not OpenState
    Next I
    
    Call frmCraftingStore.CleanInventories(False)
   
End Sub

Public Sub InitControls()
On Error GoTo ErrHandler
    
    If CustomerMode Then

        picItemsToCraft.Visible = False
    
        txtConstructionPrice.Enabled = False
        txtMaterialsPrice.Enabled = False
        imgAddRecipeToCraftButton.Visible = False
        imgOpenStoreButton.Visible = False
        
        imgCraftButton.Visible = True
        imgCraftButton.Left = 6600
        
        imgHistoryButton.Visible = False
        
    Else
        
        txtConstructionPrice.Enabled = True
        
        imgCraftButton.Visible = False
        
        imgHistoryButton.Visible = IsStoreOpen
        
        imgAddRecipeToCraftButton.Left = 6120
        imgOpenStoreButton.Left = 7560
        
        imgAddRecipeToCraftButton.Visible = True
        imgOpenStoreButton.Visible = False
        lblWorkerName.Visible = False
        
    End If
        
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub InitControls de frmCraftingStore.frm")
End Sub

Public Sub SetStoreOwnerName(ByRef StoreOwnerName As String)
     lblWorkerName.Caption = CurrentOpenStore.OwnerName
End Sub

Public Sub AddItemToInventory(ByVal ItemNumber As Integer, ByVal ConstructionPrice As Integer, ByVal Position As Integer)
    If ItemNumber <> 0 Then
        Call InvCraftableItems.SetItem(Position, ItemNumber, _
                        1, False, GameMetadata.Objs(ItemNumber).GrhIndex, _
                        0, 0, 0, 0, 0, _
                        ConstructionPrice, GameMetadata.Objs(ItemNumber).Name, 0, True)
    End If
End Sub

Public Sub DrawInventories()

    Call InvCraftableItems.DrawInventory
    frmCraftingStore.picCraftableItems.SetFocus
    
End Sub

Public Sub CleanInventories(Optional ByVal CleanItemsToCraft As Boolean = True)
    Call InvCraftableItems.DeselectItem
    Call InvCraftableItems.ClearAllSlots

    Call InvMaterialsRequired.DeselectItem
    Call InvMaterialsRequired.ClearAllSlots
    
    Call ChangeDetailsContainerVisibility(CustomerMode, True)
    
    If CleanItemsToCraft Then
        Call InvItemsToCraft.DeselectItem
        Call InvItemsToCraft.ClearAllSlots
    End If
End Sub

Public Sub CleanMetadata()

    With CurrentOpenStore
        Erase .Items
        .ItemsQty = 0
        .OwnerName = vbNullString
        .OwnerUserIndex = 0
    End With
    
    With WorkerStoreItemsToSell
        Erase .Items
        .ItemsQty = 0
        .OwnerName = vbNullString
        .OwnerUserIndex = 0
    End With

End Sub

Private Sub ImgClose_Click()
    Call CloseWindow
End Sub

Private Sub imgAddRecipeToCraftButton_Click()
    If SelectedRecipeIndex <= 0 Or SelectedRecipeNumber <= 0 Or SelectedCraftingGroup <= 0 Then Exit Sub
    
    If Not IsNumeric(txtConstructionPrice.text) Or Not IsNumeric(txtMaterialsPrice.text) Then Exit Sub
    
    If Val(txtConstructionPrice.text) = 0 Then
        Call frmMessageBox.ShowMessage("Necesitas especificar un precio de construcción.")
        Exit Sub
    End If
    
    If Not CheckMaterialsInInventory(SelectedCraftingGroup, SelectedRecipeIndex) Then
        Call frmMessageBox.ShowMessage("No tienes los materiales necesarios para construir este objeto.")
        Exit Sub
    End If
    
    Call AddRecipeToCraftingList(SelectedCraftingGroup, SelectedRecipeIndex, SelectedRecipeNumber, SelectedObjNumber, CLng(txtConstructionPrice.text), CLng(txtMaterialsPrice.text))
End Sub


Private Function CheckMaterialsInInventory(ByVal CraftingGroup As Byte, ByVal RecipeIndex As Integer) As Boolean
    
    Dim I As Integer, J As Integer
    Dim ItemQtyFound As Long

    Dim ObjList As tCurrentOpenStore
    With PlayerData.CraftingRecipeGroups(CraftingGroup).Recipes(RecipeIndex)
        For I = 1 To .MaterialsQty
            ItemQtyFound = 0
            
            For J = 1 To Inventario.MaxObjs
                If Inventario.ObjIndex(J) = .Materials(I).ObjNumber Then
                    ItemQtyFound = ItemQtyFound + Inventario.Amount(J)
                End If
            Next J
            
            If ItemQtyFound < .Materials(I).Amount Then
                CheckMaterialsInInventory = False
                Exit Function
            End If
            
        Next I
    End With
    
    CheckMaterialsInInventory = True

End Function

Private Sub imgHistoryButton_Click()
    Call frmCraftingStore_History.Show(vbModeless, frmCraftingStore)
End Sub


Private Sub imgCraftButton_Click()
    If SelectedObjNumber <= 0 Or SelectedRecipeNumber <= 0 Or SelectedRecipeIndex <= 0 Then Exit Sub
    
    Call WriteWorkerStore_CraftItem(SelectedRecipeIndex, StoreInstanceId)
End Sub

Private Sub imgOpenStoreButton_Click()

    Call WriteWorkerStore_Create(WorkerStoreItemsToSell)
    
End Sub

Private Sub picCraftableItems_Click()
        
    Call InvMaterialsRequired.DeselectItem
    Call InvMaterialsRequired.ClearAllSlots
    
    txtConstructionPrice.Visible = True
    txtMaterialsPrice.Visible = True
    lblItemName.Visible = True

    txtConstructionPrice.text = 0
    txtMaterialsPrice.text = 0
    
    If InvCraftableItems.SelectedItem < 1 Then
        Call SetFormControlsState(False, CustomerMode)
        Exit Sub
    End If
        
    Call ChangeDetailsContainerVisibility(CustomerMode, True)
    
    imgCraftButton.Enabled = True
        
    Dim I As Integer
    
    Dim IteratingObjNumber As Integer

    TempCraftingRecipeGroups = PlayerData.CraftingRecipeGroups
    TempCraftingRecipeGroupsQty = PlayerData.CraftingRecipeGroupsQty
    
    Dim ClickOnValidItem As Boolean
    
    
    ClickOnValidItem = True

    If TempCraftingRecipeGroupsQty <= 0 Then
        ClickOnValidItem = False
    End If
    If TempCraftingRecipeGroups(SelectedCraftingGroup).RecipesQty <= 0 Or InvCraftableItems.SelectedItem > TempCraftingRecipeGroups(SelectedCraftingGroup).RecipesQty Then
        ClickOnValidItem = False
    End If
    
    If Not ClickOnValidItem Then
        InvCraftableItems.DeselectItem
        Call SetFormControlsState(False, CustomerMode)
        Exit Sub
    End If
    
    SelectedObjNumber = TempCraftingRecipeGroups(SelectedCraftingGroup).Recipes(InvCraftableItems.SelectedItem).ObjNumber
    SelectedRecipeNumber = TempCraftingRecipeGroups(SelectedCraftingGroup).Recipes(InvCraftableItems.SelectedItem).RecipeIndex
    SelectedRecipeIndex = InvCraftableItems.SelectedItem
    
    lblItemName.Caption = GameMetadata.Objs(SelectedObjNumber).Name
    For I = 1 To TempCraftingRecipeGroups(SelectedCraftingGroup).Recipes(InvCraftableItems.SelectedItem).MaterialsQty
        With TempCraftingRecipeGroups(SelectedCraftingGroup).Recipes(InvCraftableItems.SelectedItem).Materials(I)
            
            If .ObjNumber <> 0 Then
    
                Call InvMaterialsRequired.SetItem(I, .ObjNumber, _
                    .Amount, False, GameMetadata.Objs(.ObjNumber).GrhIndex, _
                    0, 0, 0, 0, 0, _
                    0, GameMetadata.Objs(.ObjNumber).Name, 0, True)
   
            End If
            
        End With
    Next I
    
    ' If the store is being viewed by the a customer, then selecting an item should display the prices received by the server
    If CustomerMode Then
        txtConstructionPrice.text = TempCraftingRecipeGroups(SelectedCraftingGroup).Recipes(InvCraftableItems.SelectedItem).ConstructionPrice
        txtMaterialsPrice.text = TempCraftingRecipeGroups(SelectedCraftingGroup).Recipes(InvCraftableItems.SelectedItem).MaterialsPrice
        Call InvMaterialsRequired.DrawInventory
        Exit Sub
    End If

    ' Check if the item already exists in the list of items added to the store
    ' and show the price that was previously set.
    Dim ItemToSelIndex As Integer
    For I = 1 To WorkerStoreItemsToSell.ItemsQty
        If WorkerStoreItemsToSell.Items(I).ItemNumber = SelectedObjNumber Then
            txtConstructionPrice.text = WorkerStoreItemsToSell.Items(I).ConstructionPrice
            txtMaterialsPrice.text = WorkerStoreItemsToSell.Items(I).MaterialsPrice
            ItemToSelIndex = I
            Exit For
        End If
    Next I
    
    
    Call SetFormControlsState(True, CustomerMode)
    
    If ItemToSelIndex > 0 Then
        Call InvItemsToCraft.SelectItem(I)
        Call InvItemsToCraft.DrawInventory
    Else
        Call InvItemsToCraft.DeselectItem
    End If
    
    Call InvMaterialsRequired.DrawInventory
        
End Sub

Private Sub CleanItemSelectedData(ByVal CustomerMode As Boolean)
        imgCraftButton.Enabled = False
        
        txtConstructionPrice.text = 0
        txtMaterialsPrice.text = 0
                
        lblItemName.Caption = vbNullString
        
        Call ChangeDetailsContainerVisibility(CustomerMode, False)
        
        SelectedRecipeNumber = 0
        SelectedRecipeIndex = 0
End Sub


Public Sub AddRecipeToCraftingList(ByVal RecipeGroup As Byte, ByVal RecipeIndex As Integer, ByVal RecipeNumber As Integer, ByVal ItemNumber As Integer, ByVal CraftingPrice As Long, ByVal MaterialsPrice As Long)
       
    Dim I As Integer
    Dim ItemExists As Boolean

    ' Check if the item already exists and update the crafting price if it does
    If WorkerStoreItemsToSell.ItemsQty > 0 Then
        For I = 1 To WorkerStoreItemsToSell.ItemsQty
            If WorkerStoreItemsToSell.Items(I).ItemNumber = ItemNumber And WorkerStoreItemsToSell.Items(I).RecipeNumber = RecipeNumber Then
                WorkerStoreItemsToSell.Items(I).ConstructionPrice = CraftingPrice
                WorkerStoreItemsToSell.Items(I).MaterialsPrice = MaterialsPrice
                ItemExists = True
                Exit For
            End If
        Next I
    End If
    
    If Not ItemExists Then
        WorkerStoreItemsToSell.ItemsQty = WorkerStoreItemsToSell.ItemsQty + 1
        ReDim Preserve WorkerStoreItemsToSell.Items(1 To WorkerStoreItemsToSell.ItemsQty)
            
        With WorkerStoreItemsToSell.Items(WorkerStoreItemsToSell.ItemsQty)
            .ConstructionPrice = CraftingPrice
            .MaterialsPrice = MaterialsPrice
            .RecipeNumber = RecipeNumber
            .SelectedCraftingGroup = RecipeGroup
            .RecipeIndex = RecipeIndex
            .ItemNumber = PlayerData.CraftingRecipeGroups(RecipeGroup).Recipes(RecipeIndex).ObjNumber
        End With
    
    End If
    
    Call InvItemsToCraft.DeselectItem
    Call InvItemsToCraft.ClearAllSlots

    ' Draw the inventory
    For I = 1 To WorkerStoreItemsToSell.ItemsQty
    
        ' Reuse the variable passed to the function to avoid declaring a new one
        ItemNumber = WorkerStoreItemsToSell.Items(I).ItemNumber
        
        If WorkerStoreItemsToSell.Items(I).ItemNumber > 0 Then
             Call InvItemsToCraft.SetItem(I, ItemNumber, _
                    1, False, GameMetadata.Objs(ItemNumber).GrhIndex, _
                    0, 0, 0, 0, 0, _
                    WorkerStoreItemsToSell.Items(I).ConstructionPrice, GameMetadata.Objs(ItemNumber).Name, 0, True)
            
        End If
    Next I
    
    If WorkerStoreItemsToSell.ItemsQty > 0 Then imgOpenStoreButton.Visible = True
        
    
    Call InvItemsToCraft.DrawInventory
    
    
End Sub


Private Sub picItemsToCraft_Click()

    If InvItemsToCraft.SelectedItem = 0 Then
        txtConstructionPrice.text = 0
        txtMaterialsPrice.text = 0
        lblItemName.Caption = vbNullString
        Call SetFormControlsState(False, CustomerMode)
        InvMaterialsRequired.ClearAllSlots
        InvCraftableItems.DeselectItem
        Exit Sub
    End If
    
    txtConstructionPrice.text = WorkerStoreItemsToSell.Items(InvItemsToCraft.SelectedItem).ConstructionPrice
    txtMaterialsPrice.text = WorkerStoreItemsToSell.Items(InvItemsToCraft.SelectedItem).MaterialsPrice
    lblItemName.Caption = GameMetadata.Objs(WorkerStoreItemsToSell.Items(InvItemsToCraft.SelectedItem).ItemNumber).Name

    If Not IsStoreOpen Then
        Call RenderList(WorkerStoreItemsToSell.Items(InvItemsToCraft.SelectedItem).SelectedCraftingGroup, 1)
    End If
    
    Call InvCraftableItems.SelectItem(WorkerStoreItemsToSell.Items(InvItemsToCraft.SelectedItem).RecipeIndex)
    
    Call SetFormControlsState(True, CustomerMode)
    
End Sub


Private Sub txtConstructionPrice_Change()
    Call ValidateMaxValue(txtConstructionPrice)
End Sub

Public Sub ValidateNumericInput(ByRef KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub

Public Sub ValidateMaxValue(ByRef TxtBox As TextBox)
    On Error GoTo ErrHandler
    If Val(TxtBox.text) < 0 Then
        TxtBox.text = 1
    End If
    
    If Val(TxtBox.text) > MAX_PRICE Then
        TxtBox.text = MAX_PRICE
    End If
    
    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    TxtBox.text = MAX_PRICE
End Sub

Private Sub txtConstructionPrice_KeyPress(KeyAscii As Integer)
    Call ValidateNumericInput(KeyAscii)
End Sub

Private Sub txtMaterialsPrice_Change()
    Call ValidateMaxValue(txtMaterialsPrice)
End Sub

Private Sub txtMaterialsPrice_KeyPress(KeyAscii As Integer)
    Call ValidateNumericInput(KeyAscii)
End Sub

Public Sub ChangeDetailsContainerVisibility(ByVal CustomerMode As Boolean, ByVal Visibility As Boolean)
    
    If CustomerMode Then
        imgCraftButton.Visible = Visibility
    Else
        imgAddRecipeToCraftButton.Visible = Visibility And Not IsStoreOpen
        imgHistoryButton.Visible = Visibility And IsStoreOpen
    End If
    
    txtConstructionPrice.Visible = Visibility
    txtMaterialsPrice.Visible = Visibility
        
    lblItemName.Visible = Visibility
    
End Sub


Public Sub LoadTabButtons()
    
    Dim I As Integer
    Dim ControlStartLeft As Integer
    Dim ImgPath As String
    Dim ButtonImagePath As String
    
    Dim SeparatorWidth As Integer
    Dim AllButtonsSize As Long
    
    ImgPath = DirInterfaces & SELECTED_UI
    
    If PlayerData.CraftingRecipeGroupsQty <= 0 Then Exit Sub
    
    ReDim PicTabs(1 To PlayerData.CraftingRecipeGroupsQty)
    ReDim cPicTabs(1 To PlayerData.CraftingRecipeGroupsQty)
    
    ControlStartLeft = 0
    Dim TmpImage As Image
    
    Set TabEventHandler = New clsButtonEventHandler
    
    For I = 1 To PlayerData.CraftingRecipeGroupsQty
        
        Set PicTabs(I) = Controls.Add("VB.Image", "pictab_" & I)
        If Not (VarType(PicTabs(I)) <> vbObject) Then Load PicTabs(I)
        
        ButtonImagePath = ImgPath & PlayerData.CraftingRecipeGroups(I).TabImage & ".jpg"
        If FileExist(ButtonImagePath, vbArchive) Then
            
            Set cPicTabs(I) = New clsGraphicalButton
            Call cPicTabs(I).Initialize(PicTabs(I), ButtonImagePath, ButtonImagePath, ButtonImagePath, Me, ButtonImagePath)
            Set cPicTabs(I).EventHandler = TabEventHandler
            cPicTabs(I).Index = I
            
        End If
        PicTabs(I).Left = ControlStartLeft
        PicTabs(I).Top = 1080
        PicTabs(I).Visible = True
        
        
        AllButtonsSize = AllButtonsSize + PicTabs(I).Width
        
        ControlStartLeft = ControlStartLeft + PicTabs(I).Width + 10
    Next I

    Dim tmp As Integer
    Dim StartPosition As Integer
    StartPosition = (Me.ScaleWidth - AllButtonsSize) / 2
    SeparatorWidth = 250
    
    For I = 1 To PlayerData.CraftingRecipeGroupsQty
        PicTabs(I).Left = StartPosition
        StartPosition = StartPosition + PicTabs(I).Width + SeparatorWidth
    Next I
    
End Sub

Public Sub TabEventHandler_GraphicalButtonClick(button As clsGraphicalButton)
    SelectedCraftingGroup = button.Index
    
    Call InvCraftableItems.ClearAllSlots
    Call InvMaterialsRequired.ClearAllSlots
    Call CleanItemSelectedData(CustomerMode)
       
    Call RenderList(SelectedCraftingGroup, 1)
    
    If SelectedCraftingGroup = button.Index Then
        Call InvCraftableItems.SelectItem(1)
        Call picCraftableItems_Click
        Exit Sub
    End If

    
End Sub


Public Sub RenderList(ByVal CraftingRecipeGroup As Byte, ByVal Inicio As Integer)
On Error GoTo ErrHandler
    Dim I As Long
    Dim a As Integer
    Dim NumItems As Integer
    Dim CurrentIndex As Integer
    Dim ObjNumber As Integer
    Dim GrhIndex As Integer
    Dim Name As String

    NumItems = PlayerData.CraftingRecipeGroups(CraftingRecipeGroup).RecipesQty

    For I = 1 To MAX_ITEMS
        ' Agrego el item
        CurrentIndex = I + Inicio - 1
        
        If CurrentIndex > PlayerData.CraftingRecipeGroups(CraftingRecipeGroup).RecipesQty Then
            Exit For
        End If
        
        ObjNumber = PlayerData.CraftingRecipeGroups(CraftingRecipeGroup).Recipes(CurrentIndex).ObjNumber
        GrhIndex = GameMetadata.Objs(PlayerData.CraftingRecipeGroups(CraftingRecipeGroup).Recipes(CurrentIndex).ObjNumber).GrhIndex
        Name = GameMetadata.Objs(PlayerData.CraftingRecipeGroups(CraftingRecipeGroup).Recipes(CurrentIndex).ObjNumber).Name
  
        Call InvCraftableItems.SetItem(I, ObjNumber, 0, 0, GrhIndex, 0, 0, 0, 0, 0, 0, Name, 1)

    Next I
    Call InvCraftableItems.DrawInventory
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RenderList de frmCraft.frm")
End Sub

Public Sub SelectFirstGroup()
    If PlayerData.CraftingRecipeGroupsQty > 0 Then
        Call TabEventHandler.OnGraphicalButtonClick(cPicTabs(1))
        Call picCraftableItems_Click
    End If
End Sub

Public Sub SetFormControlsState(ByVal ValidItemSelected As Boolean, ByVal CustomerMode As Boolean)

    If ValidItemSelected Then
        txtConstructionPrice.Enabled = Not CustomerMode
        txtMaterialsPrice.Enabled = Not CustomerMode
        imgCraftButton.Visible = CustomerMode
        imgAddRecipeToCraftButton.Visible = (Not CustomerMode And Not IsStoreOpen)
    Else
        lblItemName.Caption = vbNullString
        txtConstructionPrice.Enabled = False
        txtMaterialsPrice.Enabled = False
        imgCraftButton.Visible = False
        imgAddRecipeToCraftButton.Visible = (Not CustomerMode And Not IsStoreOpen)
    End If
    
    imgOpenStoreButton.Visible = Not IsStoreOpen And Not CustomerMode
    imgHistoryButton.Visible = IsStoreOpen And Not CustomerMode
    

End Sub
