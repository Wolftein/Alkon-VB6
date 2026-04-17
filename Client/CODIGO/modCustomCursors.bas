Attribute VB_Name = "modCustomCursors"
Option Explicit

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
 
Public Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type
 
Public Enum ePictureFromFileType
    AS_BITMAP = &H0
    AS_ICON = &H1
    AS_CURSOR = &H2
End Enum

Public Enum eMousePointerPackType
    Default = 1
    Deuteranopia = 2
    Protanopia = 3
    Tripanopia = 4
    Custom = 10
    
    LastElement
End Enum


Public Enum eMousePointerModifier
    Normal = 1
    Disabled
    
    LastElement
End Enum

Public Enum eMousePointerAction
    Default = 1
    Taming
    Lumberjacking
    Mining
    Fishing
    Spell
    Arrow
    
    LastElement
End Enum

Private Const LR_LOADFROMFILE As Integer = &H10
Private CustomMousePointers(eMousePointerAction.LastElement - 1, eMousePointerModifier.LastElement - 1, 1 To eMousePointerPackType.LastElement - 1) As Picture

Public MousePointerPackNames(1 To eMousePointerPackType.LastElement - 1) As String


Private Function SetMousePointerPackNames()
    MousePointerPackNames(eMousePointerPackType.Default) = "Alkon"
    MousePointerPackNames(eMousePointerPackType.Deuteranopia) = "Deuteranopia"
    MousePointerPackNames(eMousePointerPackType.Protanopia) = "Protanopia"
    MousePointerPackNames(eMousePointerPackType.Tripanopia) = "Tripanopia"
    MousePointerPackNames(eMousePointerPackType.Custom) = "Custom"
End Function

Public Sub LoadCustomMousePointers()
    Dim CursorsPath As String
    CursorsPath = DirExtras()
    
    Call SetMousePointerPackNames

    ' Load default pointers
    Set CustomMousePointers(eMousePointerAction.Default, eMousePointerModifier.Normal, eMousePointerPackType.Default) = PictureFromFile(CursorsPath & "Default.cur", ePictureFromFileType.AS_CURSOR)
    'Set CustomMousePointers(eMousePointerAction.Default, eMousePointerModifier.Disabled) = PictureFromFile(CursorsPath & "Hand.ico", ePictureFromFileType.AS_CURSOR)
    
    Set CustomMousePointers(eMousePointerAction.Spell, eMousePointerModifier.Normal, eMousePointerPackType.Default) = PictureFromFile(CursorsPath & "Magic_Normal.cur", ePictureFromFileType.AS_CURSOR)
    Set CustomMousePointers(eMousePointerAction.Spell, eMousePointerModifier.Disabled, eMousePointerPackType.Default) = PictureFromFile(CursorsPath & "Magic_Disabled.cur", ePictureFromFileType.AS_CURSOR)
    
        
    ' Load Custom pointers. If the custom files are not present in the filesystem, then they will revert to the default ones
    Set CustomMousePointers(eMousePointerAction.Default, eMousePointerModifier.Normal, eMousePointerPackType.Custom) = GetPointer(CursorsPath, "Default", eMousePointerAction.Default, eMousePointerModifier.Normal, eMousePointerPackType.Custom)
    Set CustomMousePointers(eMousePointerAction.Spell, eMousePointerModifier.Normal, eMousePointerPackType.Custom) = GetPointer(CursorsPath, "Magic_Normal", eMousePointerAction.Spell, eMousePointerModifier.Normal, eMousePointerPackType.Custom)
    Set CustomMousePointers(eMousePointerAction.Spell, eMousePointerModifier.Disabled, eMousePointerPackType.Custom) = GetPointer(CursorsPath, "Magic_Disabled", eMousePointerAction.Spell, eMousePointerModifier.Disabled, eMousePointerPackType.Custom)
    
End Sub

Private Function GetPointer(ByRef CursorsPath As String, ByRef PointerName As String, ByVal PointerAction As eMousePointerAction, ByVal Modifier As eMousePointerModifier, ByVal MousePointerpack As eMousePointerPackType) As Picture
    ' Load custom pointers
    Dim FileName As String
    
    
    FileName = PointerName & "_" & MousePointerPackNames(MousePointerpack) & ".cur"
    
    If FileExist(CursorsPath & FileName, vbNormal) Then
        Set GetPointer = PictureFromFile(CursorsPath & FileName, ePictureFromFileType.AS_CURSOR)
    Else
        Set GetPointer = CustomMousePointers(PointerAction, Modifier, eMousePointerPackType.Default)
    End If
End Function


Public Function GetDefaultMousePointer() As Picture
    Set GetDefaultMousePointer = CustomMousePointers(eMousePointerAction.Default, eMousePointerModifier.Normal, GameConfig.Extras.MouseCursorSetToUse)
End Function

Public Function GetMousePointerForAction(ByVal Action As eMousePointerAction, ByVal Modifier As eMousePointerModifier) As Picture
    
    Set GetMousePointerForAction = CustomMousePointers(Action, Modifier, GameConfig.Extras.MouseCursorSetToUse)

    ' If the requested mouse pointer doesnt exist, the use the normal/default one.
    If GetMousePointerForAction Is Nothing Then
        Set GetMousePointerForAction = CustomMousePointers(eMousePointerAction.Default, eMousePointerModifier.Normal, GameConfig.Extras.MouseCursorSetToUse)
    End If
    
End Function

Public Sub SetFormCursorDefault(ByRef Form As Form)
    Form.MousePointer = MousePointerConstants.vbCustom
    Form.MouseIcon = CustomMousePointers(eMousePointerAction.Default, eMousePointerModifier.Normal, GameConfig.Extras.MouseCursorSetToUse)
End Sub


Private Function PictureFromFile(ByRef ImagePath As String, ByVal PictureType As ePictureFromFileType) As IPicture

Dim PicInfo As PicBmp
Dim hImage As Long
Dim tmpPic As IPictureDisp
Dim IID_IDispatch As GUID
Dim OlePictureType As PictureTypeConstants

If PictureType = AS_BITMAP Then
    OlePictureType = vbPicTypeBitmap
Else
    OlePictureType = vbPicTypeIcon
End If

    hImage = LoadImage(ByVal 0, ImagePath, PictureType, 0, 0, LR_LOADFROMFILE)

    'Setup the Guid for the function
    With IID_IDispatch
        
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    
    End With
    
    'Setup the pic structure
    With PicInfo
    
    .Size = Len(PicInfo)
        .Type = OlePictureType
        .hBmp = hImage
 
    End With
 
    'create the picture
    Call OleCreatePictureIndirect(PicInfo, IID_IDispatch, 1, tmpPic)
 
    Set PictureFromFile = tmpPic
 
End Function
