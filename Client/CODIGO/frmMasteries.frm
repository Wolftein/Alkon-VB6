VERSION 5.00
Begin VB.Form frmMasteries 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Maestrías"
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ARGENTUM.AOPictureBox PicMasteryGroupStatus 
      Height          =   480
      Left            =   1170
      TabIndex        =   4
      Top             =   1440
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   847
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   1380
      Left            =   3900
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2400
      Width           =   2430
   End
   Begin ARGENTUM.AOPictureBox PicMasteryGroupHabilities 
      Height          =   480
      Left            =   1170
      TabIndex        =   5
      Top             =   2385
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   847
   End
   Begin ARGENTUM.AOPictureBox PicMasteryGroupObjects 
      Height          =   480
      Left            =   1170
      TabIndex        =   6
      Top             =   3345
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   847
   End
   Begin VB.Label lblGoldRequired 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   1860
      Width           =   855
   End
   Begin VB.Label lblPointsRequired 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   1860
      Width           =   855
   End
   Begin VB.Label lblMasteryName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de Maestría"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   990
      Width           =   2550
   End
   Begin VB.Image imgCancelButton 
      Height          =   570
      Left            =   2520
      Top             =   4455
      Width           =   1230
   End
   Begin VB.Image imgAssignButton 
      Height          =   570
      Left            =   3870
      Top             =   4455
      Width           =   1230
   End
End
Attribute VB_Name = "frmMasteries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private SelectedMasteryGroup As Integer
Private SelectedMasteryIndex As Integer

Private MasterySlotsDevices() As Long

Private CurrentMasteriesFromGroups() As Integer

Private InvMasteryGroups() As clsGraphicalInventory
Private WithEvents GraphicalInventoryEventHandler As clsButtonEventHandler
Attribute GraphicalInventoryEventHandler.VB_VarHelpID = -1

Private Const SELECTOR_GRHINDEX As Integer = 8554
Private Const MAX_MASTERY_GROUPS As Integer = 12
Private Const MAX_MASTERY_CATEGORIES As Byte = 3

Private FormLoaded As Boolean

Private cPicAsignar As clsGraphicalButton
Private cPicCancelar As clsGraphicalButton
Public LastButtonPressed As clsGraphicalButton

Private clsForm As clsFormMovementManager


Public Property Get IsLoaded() As Boolean
    IsLoaded = FormLoaded
End Property
 
Public Property Let IsLoaded(ByVal Loaded As Boolean)
    FormLoaded = Loaded
End Property


Private Sub Form_Load()
    FormLoaded = True
    ' Defaults the selection to nothing
    SelectedMasteryIndex = 0
    
    Call LoadControls
        
    ' Initializes all the wGL devices so we can draw on them
    Call InitializeDevices(MAX_MASTERY_CATEGORIES)
    
    Call DrawAllMasteries
    
    Call modCustomCursors.SetFormCursorDefault(Me)
    
End Sub

Public Sub LoadControls()
    Dim ImgPath As String
    
    Set clsForm = New clsFormMovementManager
    clsForm.Initialize Me
    
    Set LastButtonPressed = New clsGraphicalButton
    Set cPicAsignar = New clsGraphicalButton
    Set cPicCancelar = New clsGraphicalButton
    
    ImgPath = DirInterfaces & SELECTED_UI
    Set Me.Picture = LoadPicture(ImgPath & "VentanaMaestrias.jpg")
    
    Call cPicAsignar.Initialize(imgAssignButton, ImgPath & "BotonAprender.jpg", ImgPath & "BotonAprender.jpg", ImgPath & "BotonAprender.jpg", Me, ImgPath & "BotonAprender.jpg")
    Call cPicCancelar.Initialize(imgCancelButton, ImgPath & "BotonCancelar.jpg", ImgPath & "BotonCancelar.jpg", ImgPath & "BotonCancelar.jpg", Me, ImgPath & "BotonCancelar.jpg")
    
    
End Sub

Public Sub DrawAllMasteries()
    Dim I As Integer
    Dim ClassMasteryId As Integer
  
    For I = 1 To 3
        Call LoadMasteryCategory(I)
    Next I
    
    lblMasteryName.Caption = vbNullString
    txtDescription.text = vbNullString
    lblPointsRequired.Caption = vbNullString
    lblGoldRequired.Caption = vbNullString
End Sub


Private Sub LoadMasteryCategory(ByVal CategoryId As Byte)
    Dim I As Byte
    Dim GroupStartsAt As Integer
    Dim GroupStartAt2 As Integer
    Dim ClassMasteryId As Integer
    GroupStartsAt = (4 * (CategoryId - 1))
    
    Call UIBegin(MasterySlotsDevices(CategoryId), ScaleWidth, ScaleHeight, -1)
    
    InvMasteryGroups(CategoryId).ClearAllSlots
    InvMasteryGroups(CategoryId).DeselectItem
    For I = 1 To 4
        GroupStartAt2 = GroupStartsAt + I
        If GroupStartsAt <= PlayerData.MasteryGroupsQty Then
            If PlayerData.MasteryGroups(GroupStartAt2).MasteriesQty > 0 Then
                ClassMasteryId = PlayerData.MasteryGroups(GroupStartAt2).Masteries(1)
                If PlayerData.MasteryGroups(GroupStartAt2).Masteries(1) > 0 Then
                    Call AddMasteryToInventory(CategoryId, I, ClassMasteryId, GameMetadata.Masteries(ClassMasteryId).IconGrh, GameMetadata.Masteries(ClassMasteryId).Name)
                End If
            End If
        End If
        DoEvents
    Next I
    Call UIEnd

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormLoaded = False
    Call DestroyDevices(MAX_MASTERY_CATEGORIES)
End Sub

Private Sub imgAssignButton_Click()
    
    If (SelectedMasteryIndex < 1 Or SelectedMasteryIndex > GameMetadata.MasteriesQty) Or SelectedMasteryGroup > MAX_MASTERY_GROUPS Then
        Call MsgBox("Seleccione una maestría válida")
        Exit Sub
    End If
       
    Call WriteAssignMastery(SelectedMasteryGroup, SelectedMasteryIndex)
    
End Sub

Private Sub imgAssignButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgCancelButton_Click()
    Call CloseWindow
End Sub

Private Sub CloseWindow()
On Error GoTo ErrHandler
    
    Me.Visible = False
    If frmMain.Visible Then frmMain.SetFocus
    
    Exit Sub
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CerrarVentana de frmCraft.frm")
End Sub

Private Sub imgCancelButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub


Public Sub AddMasteryToInventory(ByVal MasteryCategory As Byte, ByVal SlotIndex As Integer, ByVal MasteryId As Integer, ByVal GrhIndex As Integer, ByRef MasteryName As String)
    
    Call InvMasteryGroups(MasteryCategory).SetItem(SlotIndex, MasteryId, _
                    1, False, GrhIndex, _
                    0, 0, 0, 0, 0, _
                    0, MasteryName, 0, True)

End Sub

Public Sub DrawMasteryImage(ByVal MasterySlot As Integer, ByVal GrhIndex As Integer, ByVal ScaleWidth As Integer, ByVal ScaleHeight As Integer, ByVal LastMastery As Boolean, ByVal DrawSelection As Boolean)


    
    If LastMastery Then
        Call DrawGrhIndex(GrhIndex, 0, 0, -1#, 0, &HFF262626)
    Else
        Call DrawGrhIndex(GrhIndex, 0, 0, -1#, 0, &HFFFFFFFF)
    End If
    
    If DrawSelection Then
        Call DrawGrhIndex(SELECTOR_GRHINDEX, 0, 0, 0#, 0, &HFFFFFFFF)
    End If
End Sub

Public Sub InitializeDevices(ByVal MasteryCategories As Integer)
    Dim I As Integer
    ReDim MasterySlotsDevices(1 To MasteryCategories)
    
    'For I = 1 To DevicesQty
    '    MasterySlotsDevices(I) = Aurora_Graphics.CreatePassFromDisplay(picMastery(I).hwnd, picMastery(I).ScaleWidth, picMastery(I).ScaleHeight)
    'Next I
    
    
    Dim MaxSlots As Byte
    MaxSlots = 4
    ReDim InvMasteryGroups(1 To MasteryCategories)
    
    Set GraphicalInventoryEventHandler = New clsButtonEventHandler
    
    MasterySlotsDevices(1) = Aurora_Graphics.CreatePassFromDisplay(frmMasteries.PicMasteryGroupStatus.hwnd, frmMasteries.PicMasteryGroupStatus.ScaleWidth, frmMasteries.PicMasteryGroupStatus.ScaleHeight)
    Set InvMasteryGroups(1) = New clsGraphicalInventory
    Call InvMasteryGroups(1).Initialize(frmMasteries.PicMasteryGroupStatus, MaxSlots, , , , 10, , , , , False)
    Set InvMasteryGroups(1).EventHandler = GraphicalInventoryEventHandler
    InvMasteryGroups(1).Index = 1
    
    MasterySlotsDevices(2) = Aurora_Graphics.CreatePassFromDisplay(frmMasteries.PicMasteryGroupHabilities.hwnd, frmMasteries.PicMasteryGroupHabilities.ScaleWidth, frmMasteries.PicMasteryGroupHabilities.ScaleHeight)
    Set InvMasteryGroups(2) = New clsGraphicalInventory
    Call InvMasteryGroups(2).Initialize(frmMasteries.PicMasteryGroupHabilities, MaxSlots, , , , 10, , , , , False)
    Set InvMasteryGroups(2).EventHandler = GraphicalInventoryEventHandler
    InvMasteryGroups(2).Index = 2
    
    MasterySlotsDevices(3) = Aurora_Graphics.CreatePassFromDisplay(frmMasteries.PicMasteryGroupObjects.hwnd, frmMasteries.PicMasteryGroupObjects.ScaleWidth, frmMasteries.PicMasteryGroupObjects.ScaleHeight)
    Set InvMasteryGroups(3) = New clsGraphicalInventory
    Call InvMasteryGroups(3).Initialize(frmMasteries.PicMasteryGroupObjects, MaxSlots, , , , 10, , , , , False)
    Set InvMasteryGroups(3).EventHandler = GraphicalInventoryEventHandler
    InvMasteryGroups(3).Index = 3
    
    InvMasteryGroups(1).ClearAllSlots
    InvMasteryGroups(2).ClearAllSlots
    InvMasteryGroups(3).ClearAllSlots
    
End Sub

Public Sub GraphicalInventoryEventHandler_GraphicalInventoryClick(inventory As clsGraphicalInventory)
    Dim I As Integer
    Dim SelectedInventory As clsGraphicalInventory

    For I = 1 To 3
        If I <> inventory.Index Then
            InvMasteryGroups(I).DeselectItem
        End If
    Next I
    
    If inventory.Index < 1 Then Exit Sub
    If InvMasteryGroups(inventory.Index).SelectedItem <= 0 Then Exit Sub
    
    
    Dim CurrentSelectedElement As Integer
    CurrentSelectedElement = (4 * (inventory.Index - 1)) + InvMasteryGroups(inventory.Index).SelectedItem
    
    If PlayerData.MasteryGroupsQty <= 0 Then Exit Sub
    If CurrentSelectedElement > PlayerData.MasteryGroupsQty Then Exit Sub
    
    If PlayerData.MasteryGroups(CurrentSelectedElement).MasteriesQty <= 0 Then Exit Sub
    If PlayerData.MasteryGroups(CurrentSelectedElement).Masteries(1) = 0 Then Exit Sub
        
    SelectedMasteryGroup = CurrentSelectedElement
    SelectedMasteryIndex = PlayerData.MasteryGroups(CurrentSelectedElement).Masteries(1)
       
    If SelectedMasteryIndex <> 0 Then
        lblMasteryName.Caption = GameMetadata.Masteries(SelectedMasteryIndex).Name
        txtDescription.text = GameMetadata.Masteries(SelectedMasteryIndex).Description
        lblPointsRequired.Caption = GameMetadata.Masteries(SelectedMasteryIndex).RequiredPoints
        lblGoldRequired.Caption = GameMetadata.Masteries(SelectedMasteryIndex).RequiredGold
    End If
    
    

End Sub

Public Sub DestroyDevices(ByVal DevicesQty As Integer)
    Dim I As Integer
    
    For I = 1 To DevicesQty
        Call Aurora_Graphics.DeletePass(MasterySlotsDevices(I))
    Next I

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CloseWindow
End Sub

