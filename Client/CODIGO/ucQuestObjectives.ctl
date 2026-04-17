VERSION 5.00
Begin VB.UserControl ucQuestObjectives 
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   ScaleHeight     =   322
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   Begin ARGENTUM.AOPictureBox Pic 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _extentx        =   7011
      _extenty        =   7646
   End
End
Attribute VB_Name = "ucQuestObjectives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("UserControl")
Option Explicit

Private Const HAND_FILE_NAME As String = "pan-16.png"
Private Const SWORD_FILE_NAME As String = "sword-32.png"
Private Const TALK_FILE_NAME As String = "talk.png"
Private Const MINIATURE_NOT_FOUND As String = "NotFound.PNG"
Private Const MINIATURE_SECOND_ICON_KILL_CIRCLE As String = "IconKillTarget.PNG"


Public Enum eUserKillType
    Neutral
    Legion
    Army
    Citizen
    criminal
End Enum

Private Enum eQuestRequirementType
    CollectObject
    KillNPC
    KillUser
    TalkNpc
End Enum

Private Type tControlConfiguration
    TextOffsetX As Integer
    TextOffsetY As Integer
    MarginX As Integer
    MarginY As Integer
    ControlFont As tFuente
    SpaceBetweenItemAndText As Integer
    SpaceBetweenItems As Integer
End Type

Private ControlConfiguration As tControlConfiguration

Private Type tItemData
    ItemIndex As Long
    GrhIndex As Long
    NpcMiniatureFilePath As String
    Quantity As Long
    RequiredQuantity As Long
    Description As String
    RequirementType As eQuestRequirementType
End Type

Private Const SUCCESS_FONT_COLOR As Long = &HFF00FF00

Private ItemList(200) As tItemData
Private ItemListCount As Integer

Private HandMaterial As Integer
Private SwordMaterial As Integer
Private TalkMaterial As Integer
Private KillCircleMaterial As Integer

Private Device As Long

Private MustInvalidate As Boolean

Private MiniaturesPath As String



Public Sub Initialize()
    Device = Aurora_Graphics.CreatePassFromDisplay(Pic.hwnd, Pic.Width, Pic.Height)
    
    With ControlConfiguration
        Set .ControlFont.Asset = FuentesJuego.Inventarios.Asset
        .ControlFont.color = FuentesJuego.Inventarios.color
        .ControlFont.Tamanio = FuentesJuego.Inventarios.Tamanio
        .SpaceBetweenItemAndText = 40
        .SpaceBetweenItems = 40
        .MarginX = 10
        
    End With
    ItemListCount = 0
    
    Dim GrhPath As String
    
    GrhPath = DirInterfaces & SELECTED_UI
    GrhPath = Replace(GrhPath, "\\", "\")
    
    MiniaturesPath = DirNpcMiniatures()
    
    HandMaterial = CreateMaterialWithTextureFromFile(GrhPath & HAND_FILE_NAME)
    SwordMaterial = CreateMaterialWithTextureFromFile(GrhPath & SWORD_FILE_NAME)
    TalkMaterial = CreateMaterialWithTextureFromFile(GrhPath & TALK_FILE_NAME)
    
    KillCircleMaterial = CreateMaterialWithTextureFromFile(MiniaturesPath & MINIATURE_SECOND_ICON_KILL_CIRCLE)

    Call EnsureInvalidate
End Sub
Public Sub Clear()
    ItemListCount = 0
End Sub
Public Sub SetUserKill(ByVal KillType As eUserKillType, ByVal Quantity As Long, ByVal RequiredQuantity As Long)
    Dim Index As Long
    For Index = 0 To ItemListCount - 1
         With ItemList(Index)
            If .ItemIndex = KillType And .RequirementType = eQuestRequirementType.KillUser Then
                .Quantity = Quantity
                .RequiredQuantity = RequiredQuantity
                
                Call EnsureInvalidate
                
                Exit Sub
            End If
        End With
    Next Index
    
    With ItemList(Index)
        .RequirementType = eQuestRequirementType.KillUser
        .ItemIndex = KillType
        .Quantity = Quantity
        .RequiredQuantity = RequiredQuantity
        .Description = "Matar " & KillType
    End With
    
    ItemListCount = ItemListCount + 1
    
    Call EnsureInvalidate
End Sub
Public Sub SetTalk(ByVal NpcIndex As Long, ByVal AlreadyDone As Boolean)
    Dim Index As Long
    Dim Quantity As Long
    Dim RequiredQuantity As Long
    
    RequiredQuantity = 1
    
    If AlreadyDone Then
        Quantity = 1
    Else
        Quantity = 0
    End If
    For Index = 0 To ItemListCount - 1
         With ItemList(Index)
            If .ItemIndex = NpcIndex And .RequirementType = eQuestRequirementType.TalkNpc Then
                .Quantity = Quantity
                .RequiredQuantity = RequiredQuantity
                
                Call EnsureInvalidate
                
                Exit Sub
            End If
        End With
    Next Index
    
    With ItemList(Index)
        .RequirementType = eQuestRequirementType.TalkNpc
        .ItemIndex = NpcIndex
        .GrhIndex = 697
        .Quantity = Quantity
        .RequiredQuantity = RequiredQuantity
        .Description = GameMetadata.Npcs(NpcIndex).Name
    End With
    
    ItemListCount = ItemListCount + 1
    
    Call EnsureInvalidate
End Sub

Public Sub SetNpcKill(ByVal NpcIndex As Long, ByVal Quantity As Long, ByVal RequiredQuantity As Long, ByRef NpcMiniatureFileName As String)
    Dim Index As Long
    For Index = 0 To ItemListCount - 1
         With ItemList(Index)
            If .ItemIndex = NpcIndex And .RequirementType = eQuestRequirementType.KillNPC Then
                .Quantity = Quantity
                .RequiredQuantity = RequiredQuantity
                
                Call EnsureInvalidate
                
                Exit Sub
            End If
        End With
    Next Index
    
    With ItemList(Index)
        .RequirementType = eQuestRequirementType.KillNPC
        .ItemIndex = NpcIndex
        .GrhIndex = 697
        .Quantity = Quantity
        .RequiredQuantity = RequiredQuantity
        .Description = "Matar " & GameMetadata.Npcs(NpcIndex).Name
        .NpcMiniatureFilePath = GetMiniFilePathOrNotFound(MiniaturesPath, NpcMiniatureFileName)
    End With
    
    ItemListCount = ItemListCount + 1
    
    Call EnsureInvalidate
End Sub
Public Sub SetItem(ByVal ObjectIndex As Long, ByVal Quantity As Long, ByVal RequiredQuantity As Long)
    Dim Index As Long
    For Index = 0 To ItemListCount - 1
         With ItemList(Index)
            If .ItemIndex = ObjectIndex And .RequirementType = eQuestRequirementType.CollectObject Then
                .Quantity = Quantity
                .RequiredQuantity = RequiredQuantity
                
                Call EnsureInvalidate
                
                Exit Sub
            End If
        End With
    Next Index
    
    With ItemList(Index)
        .RequirementType = eQuestRequirementType.CollectObject
        .ItemIndex = ObjectIndex
        .GrhIndex = GameMetadata.Objs(ObjectIndex).GrhIndex
        .Quantity = Quantity
        .RequiredQuantity = RequiredQuantity
        .Description = GameMetadata.Objs(ObjectIndex).Name
    End With
    
    ItemListCount = ItemListCount + 1
    Call EnsureInvalidate
    
End Sub

Private Function GetMiniFilePathOrNotFound(ByRef Folder As String, ByRef fileName As String) As String
    Dim PathToReturn  As String
    
    GetMiniFilePathOrNotFound = Folder & fileName
    
    If FileExist(GetMiniFilePathOrNotFound, vbArchive) Then Exit Function
    
    GetMiniFilePathOrNotFound = MiniaturesPath & MINIATURE_NOT_FOUND

End Function

Public Sub DrawItems()
    Dim X As Integer
    Dim Y As Integer
    
    Dim SourceRect As Math_Rectf
    SourceRect.X1 = 0: SourceRect.X2 = 1
    SourceRect.Y1 = 0: SourceRect.Y2 = 1
    
    Dim DestinationRect As Math_Rectf
    Dim FontColor As Long
    
    Dim ItemData As tItemData
    
    ItemData.GrhIndex = 550
    ItemData.Quantity = 2000
    ItemData.RequiredQuantity = 2000
    
    FontColor = &HFF00FF00
    X = 0
    Y = 0
    
    Call UIBegin(Device, Pic.ScaleWidth, Pic.ScaleHeight, &H0)

    With ControlConfiguration
        Dim Index As Long
        For Index = 0 To ItemListCount - 1
            
            Y = Index * .SpaceBetweenItems
            If ItemList(Index).Quantity >= ItemList(Index).RequiredQuantity Then
                FontColor = SUCCESS_FONT_COLOR
            Else
                FontColor = .ControlFont.color
            End If
            If ItemList(Index).RequirementType = eQuestRequirementType.CollectObject Then
                DestinationRect.X1 = .MarginX + X + 16
                DestinationRect.Y1 = .MarginY + Y + 16
                DestinationRect.X2 = DestinationRect.X1 + 16
                DestinationRect.Y2 = DestinationRect.Y1 + 16
                
                Call DrawMaterial(DestinationRect, SourceRect, GetDepth(1, X, Y, 2), 0, &HFFFFFFFF, HandMaterial)
                Call DrawGrhIndex(ItemList(Index).GrhIndex, .MarginX + X, .MarginY + Y, GetDepth(1, X, Y, 1), 0, &HFFFFFFFF)
                Call DrawText(.MarginX + X + .SpaceBetweenItemAndText, .MarginY + Y + 8, GetDepth(1, X, Y, 3), ItemList(Index).Description, .ControlFont.color, eRendererAlignmentLeftMiddle, .ControlFont, True)
                Call DrawText(.MarginX + X + .SpaceBetweenItemAndText, .MarginY + Y + 24, GetDepth(1, X, Y, 3), ItemList(Index).Quantity & " / " & ItemList(Index).RequiredQuantity, FontColor, eRendererAlignmentLeftMiddle, .ControlFont, True)
            End If
            
            If ItemList(Index).RequirementType = eQuestRequirementType.KillNPC Then
                DestinationRect.X1 = .MarginX + X
                DestinationRect.Y1 = .MarginY + Y
                DestinationRect.X2 = DestinationRect.X1 + 32
                DestinationRect.Y2 = DestinationRect.Y1 + 32
                Call DrawMaterial(DestinationRect, SourceRect, GetDepth(1, X, Y, 2), 0, &HFFFFFFFF, CreateMaterialWithTextureFromFile(ItemList(Index).NpcMiniatureFilePath))
                Call DrawMaterial(DestinationRect, SourceRect, GetDepth(1, X, Y, 2), 0, &HFFFFFFFF, KillCircleMaterial)
                Call DrawText(.MarginX + X + .SpaceBetweenItemAndText, .MarginY + Y + 8, GetDepth(1, X, Y, 3), ItemList(Index).Description, .ControlFont.color, eRendererAlignmentLeftMiddle, .ControlFont, True)
                Call DrawText(.MarginX + X + .SpaceBetweenItemAndText, .MarginY + Y + 24, GetDepth(1, X, Y, 3), ItemList(Index).Quantity & " / " & ItemList(Index).RequiredQuantity, FontColor, eRendererAlignmentLeftMiddle, .ControlFont, True)
            End If
            
            If ItemList(Index).RequirementType = eQuestRequirementType.KillUser Then
                DestinationRect.X1 = .MarginX + X
                DestinationRect.Y1 = .MarginY + Y
                DestinationRect.X2 = DestinationRect.X1 + 32
                DestinationRect.Y2 = DestinationRect.Y1 + 32
                Call DrawMaterial(DestinationRect, SourceRect, GetDepth(1, X, Y, 2), 0, &HFFFFFFFF, SwordMaterial)
                Call DrawText(.MarginX + X + .SpaceBetweenItemAndText, .MarginY + Y + 8, GetDepth(1, X, Y, 3), ItemList(Index).Description, .ControlFont.color, eRendererAlignmentLeftMiddle, .ControlFont, True)
                Call DrawText(.MarginX + X + .SpaceBetweenItemAndText, .MarginY + Y + 24, GetDepth(1, X, Y, 3), ItemList(Index).Quantity & " / " & ItemList(Index).RequiredQuantity, FontColor, eRendererAlignmentLeftMiddle, .ControlFont, True)
            End If
            
            If ItemList(Index).RequirementType = eQuestRequirementType.TalkNpc Then
                DestinationRect.X1 = .MarginX + X
                DestinationRect.Y1 = .MarginY + Y
                DestinationRect.X2 = DestinationRect.X1 + 32
                DestinationRect.Y2 = DestinationRect.Y1 + 32
                Call DrawMaterial(DestinationRect, SourceRect, GetDepth(1, X, Y, 2), 0, &HFFFFFFFF, TalkMaterial)
                Call DrawText(.MarginX + X + .SpaceBetweenItemAndText, .MarginY + Y + 8, GetDepth(1, X, Y, 3), "Hablar con:", .ControlFont.color, eRendererAlignmentLeftMiddle, .ControlFont, True)
                Call DrawText(.MarginX + X + .SpaceBetweenItemAndText, .MarginY + Y + 24, GetDepth(1, X, Y, 3), ItemList(Index).Description, .ControlFont.color, eRendererAlignmentLeftMiddle, .ControlFont, True)
            End If
            
        Next Index
    End With
    
    Call UIEnd
    
    MustInvalidate = False
End Sub

Private Sub Pic_Paint()
    If Device > 0 Then
        Call DrawItems
    End If
End Sub

Private Sub UserControl_Initialize()
    Pic.Width = ScaleWidth
    Pic.Height = ScaleHeight
End Sub
Private Sub EnsureInvalidate()
    If Not MustInvalidate Then
        Call Invalidate(Pic.hwnd)
        MustInvalidate = True
    End If
End Sub
Private Sub UserControl_Terminate()
    If Device > 0 Then
        'TODO: Wolftein
        'Call wGL_Graphic_Renderer.Destroy_Material(HandMaterial)
        'Call wGL_Graphic_Renderer.Destroy_Material(SwordMaterial)
        'Call wGL_Graphic_Renderer.Destroy_Material(TalkMaterial)
        'Call wGL_Graphic_Renderer.Destroy_Material(KillCircleMaterial)
        
        Call Aurora_Graphics.DeletePass(Device)
        
        HandMaterial = 0
        SwordMaterial = 0
        TalkMaterial = 0
    End If
End Sub
