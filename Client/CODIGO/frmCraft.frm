VERSION 5.00
Begin VB.Form frmCraft 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8295
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   473
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   553
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrAutoCraft 
      Left            =   480
      Top             =   240
   End
   Begin ARGENTUM.AOPictureBox picSelItem 
      Height          =   480
      Left            =   5085
      TabIndex        =   11
      Top             =   3135
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.VScrollBar Scroll 
      Height          =   5775
      Left            =   120
      TabIndex        =   1
      Top             =   1170
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtCantItems 
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
      Height          =   255
      Left            =   6720
      MaxLength       =   5
      TabIndex        =   0
      Text            =   "1"
      Top             =   2265
      Width           =   1290
   End
   Begin VB.Label lblSkills 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4230
      TabIndex        =   10
      Top             =   5085
      Width           =   3615
   End
   Begin VB.Image imgMarcoItem 
      Height          =   570
      Index           =   7
      Left            =   480
      Top             =   6210
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   7
      Left            =   1140
      TabIndex        =   9
      Top             =   6345
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Image imgMarcoItem 
      Height          =   570
      Index           =   6
      Left            =   480
      Top             =   5374
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   6
      Left            =   1140
      TabIndex        =   8
      Top             =   5505
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Label lblSelItem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4275
      TabIndex        =   7
      Top             =   1920
      Width           =   2100
   End
   Begin VB.Label lblItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   5
      Left            =   1140
      TabIndex        =   6
      Top             =   4665
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Label lblItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   4
      Left            =   1140
      TabIndex        =   5
      Top             =   3885
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Label lblItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   3
      Left            =   1140
      TabIndex        =   4
      Top             =   3090
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Label lblItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   1140
      TabIndex        =   3
      Top             =   2295
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Label lblItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   1140
      TabIndex        =   2
      Top             =   1500
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Image imgMarcoItem 
      Height          =   570
      Index           =   5
      Left            =   480
      Top             =   4534
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image picConstruir 
      Height          =   420
      Left            =   5280
      Top             =   6120
      Width           =   1710
   End
   Begin VB.Image imgMarcoItem 
      Height          =   570
      Index           =   1
      Left            =   487
      Top             =   1395
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image picCheckBox 
      Height          =   225
      Left            =   7245
      MousePointer    =   99  'Custom
      Top             =   3420
      Width           =   225
   End
   Begin VB.Image imgCerrar 
      Height          =   240
      Left            =   7920
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgMarcoItem 
      Height          =   570
      Index           =   2
      Left            =   480
      Top             =   2167
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image imgMarcoItem 
      Height          =   570
      Index           =   3
      Left            =   480
      Top             =   2948
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image imgMarcoItem 
      Height          =   570
      Index           =   4
      Left            =   480
      Top             =   3756
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmCraft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tControlArrayPositions
    Top As Long
    Left As Long
End Type

Private Enum eStation
    Herreria
    Carpinteria
End Enum

Private Enum eSubStation
    Armas
    Armaduras
    ObjCarpinteria
End Enum

Private picCheck As Picture
Private picRecuadroItem As Picture

Private UltimaPestania As Byte

Private TabType() As Byte
Private PicTab() As VB.Image
Private cPicTab() As clsGraphicalButton
Private WithEvents TabEventHandler As clsButtonEventHandler
Attribute TabEventHandler.VB_VarHelpID = -1
Private WithEvents CraftingRecipeEventHandler As clsButtonEventHandler
Attribute CraftingRecipeEventHandler.VB_VarHelpID = -1

' Item boxes
'   Recipes
Private PicRecipes() As VBControlExtender 'AOPictureBox
Private PicMaterials() As VBControlExtender 'AOPictureBox
Private RecipesPositions() As tControlArrayPositions
Private MaterialsPositions() As tControlArrayPositions

Private cPicCerrar As clsGraphicalButton
Private cPicConstruir As clsGraphicalButton
Public LastButtonPressed As clsGraphicalButton

Private Cargando As Boolean

Private UsarMacro As Boolean

Private Station As Byte
Private SubStation As Byte
Private SelItem As Integer
Private SelectedGroup As Byte

Private MacroCraftingGroup As Byte
Private MacroCraftRecipeIndex As Integer
Private MacroCraftQty As Integer

Private clsFormulario As clsFormMovementManager
Private CraftList() As tItemsConstruibles
Private LastTickCount As Long

Private Sub CargarImagenes()
On Error GoTo ErrHandler
  
    Dim ImgPath As String
    Dim Index As Integer
    
    ImgPath = DirInterfaces & SELECTED_UI
    
    Set Me.Picture = LoadPicture(ImgPath & "VentanaCrafteo.jpg")
    
    Set picCheck = LoadPicture(ImgPath & "BotonCheck.jpg")
    
    Set picRecuadroItem = LoadPicture(ImgPath & "MarcoItemsCrafteo.jpg")
    
    For Index = 1 To MAX_LIST_ITEMS
        imgMarcoItem(Index).Picture = picRecuadroItem
    Next Index
    
    Set cPicCerrar = New clsGraphicalButton
    Set cPicConstruir = New clsGraphicalButton
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cPicCerrar.Initialize(imgCerrar, ImgPath & "BotonCrafteoSalir.jpg", ImgPath & "BotonCrafteoSalirRollover.jpg", ImgPath & "BotonCrafteoSalirClick.jpg", Me, ImgPath & "BotonCrafteoSalirDisabled.jpg")
    Call cPicConstruir.Initialize(picConstruir, ImgPath & "BotonConstruir.jpg", ImgPath & "BotonConstruirHover.jpg", ImgPath & "BotonConstruirClick.jpg", Me, ImgPath & "BotonConstruirDisabled.jpg")

    picCheckBox.MouseIcon = picMouseIcon
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CargarImagenes de frmCraft.frm")
End Sub

Public Sub SetDefaultItemsPositions()
    ReDim RecipesPositions(1 To MAX_LIST_ITEMS)
    
    RecipesPositions(1).Top = 96
    RecipesPositions(2).Top = 147
    RecipesPositions(3).Top = 200
    RecipesPositions(4).Top = 253
    RecipesPositions(5).Top = 305
    RecipesPositions(6).Top = 361
    RecipesPositions(7).Top = 417
    
    ReDim MaterialsPositions(1 To MAX_CRAFT_MATERIAL)
    MaterialsPositions(1).Top = 161
    MaterialsPositions(1).Left = 287
    MaterialsPositions(2).Top = 161
    MaterialsPositions(2).Left = 339
    MaterialsPositions(3).Top = 161
    MaterialsPositions(3).Left = 391
    MaterialsPositions(4).Top = 209
    MaterialsPositions(4).Left = 287
    MaterialsPositions(5).Top = 209
    MaterialsPositions(5).Left = 391
    MaterialsPositions(6).Top = 259
    MaterialsPositions(6).Left = 287
    MaterialsPositions(7).Top = 259
    MaterialsPositions(7).Left = 339
    MaterialsPositions(8).Top = 259
    MaterialsPositions(8).Left = 391
    
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
    
    ReDim PicTab(1 To PlayerData.CraftingRecipeGroupsQty)
    ReDim cPicTab(1 To PlayerData.CraftingRecipeGroupsQty)
    
    ControlStartLeft = 0
    Dim TmpImage As Image
    
    Set TabEventHandler = New clsButtonEventHandler
    
    For I = 1 To PlayerData.CraftingRecipeGroupsQty
        
        Set PicTab(I) = Controls.Add("VB.Image", "pictab_" & I)
        If Not (VarType(PicTab(I)) <> vbObject) Then Load PicTab(I)
        
        ButtonImagePath = ImgPath & PlayerData.CraftingRecipeGroups(I).TabImage & ".jpg"
        If FileExist(ButtonImagePath, vbArchive) Then
            
            Set cPicTab(I) = New clsGraphicalButton
            Call cPicTab(I).Initialize(PicTab(I), ButtonImagePath, ButtonImagePath, ButtonImagePath, Me, ButtonImagePath)
            Set cPicTab(I).EventHandler = TabEventHandler
            cPicTab(I).Index = I
        Else
            
        End If
        PicTab(I).Left = ControlStartLeft
        PicTab(I).Top = 20
        PicTab(I).Visible = True
        
        
        AllButtonsSize = AllButtonsSize + PicTab(I).Width
        
        ControlStartLeft = ControlStartLeft + PicTab(I).Width + 10
    Next I

    Dim tmp As Integer
    Dim StartPosition As Integer
    StartPosition = (Me.ScaleWidth - AllButtonsSize) / 2
    SeparatorWidth = 20
    
    For I = 1 To PlayerData.CraftingRecipeGroupsQty
        PicTab(I).Left = StartPosition
        
        StartPosition = StartPosition + PicTab(I).Width + SeparatorWidth
    Next I

End Sub

Private Sub ConstruirItem()
On Error GoTo ErrHandler
  
    Dim ItemIndex As Integer
    Dim CraftQty As Integer
    
    If Scroll.Visible = True Then ItemIndex = Scroll.value
    ItemIndex = PlayerData.CraftingRecipeGroups(SelectedGroup).Recipes(SelItem).RecipeIndex
    CraftQty = Val(txtCantItems.Text)
    
    Call WriteInitCrafting(CraftQty)
    Call WriteCraftItem(SelectedGroup, ItemIndex, UsarMacro)
    
    If UsarMacro Then
        MacroCraftingGroup = SelectedGroup
        MacroCraftRecipeIndex = ItemIndex
        MacroCraftQty = CraftQty
        Call EnableAutoCrafting
        
        Call cPicConstruir.EnableButton(False)
    Else
        Call CerrarVentana
    End If

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ConstruirItem de frmCraft.frm")
End Sub

Public Sub Initialize()
On Error GoTo ErrHandler
    Dim MaxConstItem As Integer
    Dim I As Integer
    Dim IsInitialized As Boolean
    Dim TempPic As AOPictureBox
        
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    LastTickCount = 0
    
    Call SetDefaultItemsPositions
    
    ' Load crafting Item containers
    ReDim PicRecipes(1 To MAX_LIST_ITEMS)
    Set CraftingRecipeEventHandler = New clsButtonEventHandler
    For I = 1 To MAX_LIST_ITEMS
        Set PicRecipes(I) = Controls.Add("Argentum.AOPictureBox", "recipe_" & I, Me)

        PicRecipes(I).Left = 35
        PicRecipes(I).Top = RecipesPositions(I).Top
        PicRecipes(I).ToolTipText = ""
        PicRecipes(I).Width = 32
        PicRecipes(I).Height = 32
        
        Set InvCraftItem(I) = New clsGraphicalInventory
        Set TempPic = PicRecipes(I)
        Call InvCraftItem(I).Initialize(TempPic, 1, , , , , , False)
        
        InvCraftItem(I).Index = I
        Set InvCraftItem(I).EventHandler = CraftingRecipeEventHandler
    Next I
    
    ' Load crafting Item containers
    ReDim PicMaterials(1 To MAX_CRAFT_MATERIAL)
    For I = 1 To MAX_CRAFT_MATERIAL
        Set PicMaterials(I) = Controls.Add("Argentum.AOPictureBox", "material_" & I, Me)

        PicMaterials(I).Left = MaterialsPositions(I).Left
        PicMaterials(I).Top = MaterialsPositions(I).Top
        PicMaterials(I).ToolTipText = ""
        PicMaterials(I).Width = 32
        PicMaterials(I).Height = 32
        
        Set TempPic = PicMaterials(I)
        Set InvCraftMaterial(I) = New clsGraphicalInventory
        Call InvCraftMaterial(I).Initialize(TempPic, 10, , , , , , False)
        
        PicMaterials(I).Visible = True
    Next I
        
    Set InvCraftSelItem = New clsGraphicalInventory
    Call InvCraftSelItem.Initialize(picSelItem, 1, , , , , , False)
 
    Call HideExtraControls(1)
    
    ' By default, when the form is initialized, we're not using the macro.
    UsarMacro = False
    picCheckBox.Picture = Nothing

    Call LoadTabButtons
    Call CargarImagenes
    
    ' Disable the craft button. It will be enabled when selecting one of the elements to craft.
    Call cPicConstruir.EnableButton(False)
    
    If PlayerData.CraftingRecipeGroupsQty > 0 Then
        SelectedGroup = 1
        SelItem = 1
        Call RenderList(1, 1)
    End If
    
    Cargando = False
    MirandoCarpinteria = True
  
  Exit Sub
  
ErrHandler:
  Call LogError("AO: Error" & Err.Number & "(" & Err.Description & ") en Sub Initialize de frmCraft.frm")
  CerrarVentana
End Sub

Public Sub HideExtraControls(ByVal CraftingGroup As Integer)
On Error GoTo ErrHandler
  
    Dim I As Integer
    Dim NumItems As Integer
    NumItems = PlayerData.CraftingRecipeGroups(CraftingGroup).RecipesQty
    
    
    For I = 1 To MAX_LIST_ITEMS
        PicRecipes(I).Visible = (NumItems >= I)
        imgMarcoItem(I).Visible = (NumItems >= I)
        lblItem(I).Visible = (NumItems >= I)
    Next I

    If NumItems > MAX_LIST_ITEMS Then
        Scroll.Visible = True
        Cargando = True
        Scroll.Max = NumItems - MAX_LIST_ITEMS
        Cargando = False
    Else
        Scroll.Visible = False
    End If
    
    lblSelItem.Caption = ""
    
    Call InvCraftSelItem.SetItem(1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0, True)
    For I = 1 To MAX_CRAFT_MATERIAL
        Call InvCraftMaterial(I).SetItem(1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0, True)
    Next I
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HideExtraControls de frmCraft.frm")
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

    For I = 1 To MAX_LIST_ITEMS
        ' Agrego el item
        CurrentIndex = I + Inicio - 1
        
        If CurrentIndex > PlayerData.CraftingRecipeGroups(CraftingRecipeGroup).RecipesQty Then
            Exit For
        End If
        
        ObjNumber = PlayerData.CraftingRecipeGroups(CraftingRecipeGroup).Recipes(CurrentIndex).ObjNumber
        GrhIndex = GameMetadata.Objs(PlayerData.CraftingRecipeGroups(CraftingRecipeGroup).Recipes(CurrentIndex).ObjNumber).GrhIndex
        Name = GameMetadata.Objs(PlayerData.CraftingRecipeGroups(CraftingRecipeGroup).Recipes(CurrentIndex).ObjNumber).Name
  
        Call InvCraftItem(I).SetItem(1, ObjNumber, 0, 0, GrhIndex, 0, 0, 0, 0, 0, 0, Name, 1)
        PicRecipes(I).ToolTipText = Name
        lblItem(I).Caption = Name
        PicRecipes(I).Visible = True
        lblItem(I).Visible = True
        imgMarcoItem(I).Visible = True
        Call InvCraftItem(I).DrawInventory
    Next I
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RenderList de frmCraft.frm")
End Sub

Public Sub RenderSelItem()
On Error GoTo ErrHandler
  
    Dim I As Long
    Dim a As Integer
    Dim NumItems As Integer
    Dim CraftList() As tItemsConstruibles
    
    If SelItem = 0 Then Exit Sub
    
    ' Disable the craft button. It will be enabled when selecting one of the elements to craft.
    Call cPicConstruir.EnableButton(SelItem <> 0)
    
    NumItems = PlayerData.CraftingRecipeGroups(SelectedGroup).Recipes(SelItem).MaterialsQty
    Dim ObjIndex As Integer
    Dim GrhIndex As Integer
    Dim Name As String
    Dim Amount As Integer
    
    ObjIndex = PlayerData.CraftingRecipeGroups(SelectedGroup).Recipes(SelItem).ObjNumber
    GrhIndex = GameMetadata.Objs(ObjIndex).GrhIndex
    Name = GameMetadata.Objs(ObjIndex).Name
    
    Call InvCraftSelItem.SetItem(1, ObjIndex, 0, 0, GrhIndex, 0, 0, 0, 0, 0, 0, Name, 0)
    picSelItem.ToolTipText = Name
    lblSelItem.Caption = Name
  
    For I = 1 To MAX_CRAFT_MATERIAL
        Call InvCraftMaterial(I).SetItem(1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0, True)
    Next I
    
    For I = 1 To PlayerData.CraftingRecipeGroups(SelectedGroup).Recipes(SelItem).MaterialsQty
        ' We will only render 8 items. If there's more, the dats are wrong.
        If I > NumItems Then Exit For
        
        ObjIndex = PlayerData.CraftingRecipeGroups(SelectedGroup).Recipes(SelItem).Materials(I).ObjNumber
        Amount = PlayerData.CraftingRecipeGroups(SelectedGroup).Recipes(SelItem).Materials(I).Amount
        GrhIndex = GameMetadata.Objs(ObjIndex).GrhIndex
        Name = GameMetadata.Objs(ObjIndex).Name

        PicMaterials(I).ToolTipText = GameMetadata.Objs(ObjIndex).Name
        Call InvCraftMaterial(I).SetItem(1, 0, Amount, 0, GrhIndex, 0, 0, 0, 0, 0, 0, Name, 0)
        Call InvCraftMaterial(I).DrawInventory
        
    Next I

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RenderSelItem de frmCraft.frm")
End Sub


Private Sub Form_Load()
    Call modCustomCursors.SetFormCursorDefault(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandler

    MirandoHerreria = False
    MirandoCarpinteria = False
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Unload de frmCraft.frm")
End Sub

Private Sub imgCerrar_Click()
    Call CerrarVentana
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CerrarVentana
End Sub

Private Sub CerrarVentana()
On Error GoTo ErrHandler

    Call DisableAutoCrafting

    Unload Me
    If frmMain.Visible Then frmMain.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CerrarVentana de frmCraft.frm")
End Sub

Private Sub picCheckBox_Click()
    
On Error GoTo ErrHandler
  
    UsarMacro = Not UsarMacro

    If UsarMacro Then
        picCheckBox.Picture = picCheck
    Else
        picCheckBox.Picture = Nothing
    End If

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub picCheckBox_Click de frmCraft.frm")
End Sub

Private Sub picConstruir_Click()
    If cPicConstruir.IsEnabled Then
        Call ConstruirItem
    End If
End Sub


Private Sub picMaterial_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub


Private Sub picSelItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub Scroll_Change()
On Error GoTo ErrHandler
  
    Dim I As Long
    Dim ActualTick As Long
    
    If Cargando Then Exit Sub
    
    I = Scroll.value

    ActualTick = GetTickCount
    
    If (ActualTick - LastTickCount > 100) Then
        LastTickCount = ActualTick
        Call RenderList(SelectedGroup, I + 1)
    End If
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Scroll_Change de frmCraft.frm")
End Sub


Private Sub tmrAutoCraft_Timer()
    If Not UsarMacro Then Exit Sub
  
    Call WriteInitCrafting(MacroCraftQty)
    Call WriteCraftItem(MacroCraftingGroup, MacroCraftRecipeIndex, True)
End Sub

Private Sub txtCantItems_Change()
On Error GoTo ErrHandler
    If Val(txtCantItems.text) < 0 Then
        txtCantItems.text = 1
    End If
    
    If Val(txtCantItems.text) > MAX_INVENTORY_OBJS Then
        txtCantItems.text = MAX_INVENTORY_OBJS
    End If
    
    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    txtCantItems.text = MAX_INVENTORY_OBJS
End Sub

Private Sub txtCantItems_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandler
  
    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub txtCantItems_KeyPress de frmCraft.frm")
End Sub

Public Sub TabEventHandler_GraphicalButtonClick(button As clsGraphicalButton)
    SelectedGroup = button.Index
    
    Call HideExtraControls(SelectedGroup)
    If Scroll.value <> 0 Then
        Scroll.value = 0
    Else
        Call RenderList(SelectedGroup, 1)
    End If
    
End Sub

Public Sub CraftingRecipeEventHandler_GraphicalInventoryClick(inventory As clsGraphicalInventory)
    SelItem = inventory.Index + Scroll.value
    Call RenderSelItem
End Sub

Public Sub EnableAutoCrafting()
    
    tmrAutoCraft.Interval = PlayerData.Intervals.WorkMacro
    tmrAutoCraft.Enabled = True
    
    
End Sub

Public Sub DisableAutoCrafting()
    tmrAutoCraft.Enabled = False
End Sub

