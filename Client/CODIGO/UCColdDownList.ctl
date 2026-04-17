VERSION 5.00
Begin VB.UserControl UCColdDownList 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer TimerRender 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2640
      Top             =   360
   End
   Begin VB.ListBox InternalList 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   720
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "UCColdDownList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type CooldownListData
    Name As String
    Cooldown As Long
    startTime As Long
    ShowOnlyDefault As Boolean
End Type

Private Const LB_GETITEMHEIGHT = &H1A1

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private CooldownList(255) As CooldownListData
Private CooldownListCount As Integer

Private DefaultCoolDown As Long
Private DefaultCooldownStartTime As Long

Private SpellAfterMeleeCoolDown As Long
Private SpellAfterMeleeCoolDownStartTime As Long

Private Brush As Long
Private BackBrush As Long

Private ItemHeight As Long
Private ListHdc As Long

Private mBarOffsetY As Integer
Private mBarColor As OLE_COLOR
Private mBarBackColor As OLE_COLOR

Public Event SpellAfterMeleeCooldownFinish()
Public Event CooldownFinish(Index As Integer)
Public Event SelectionChanged()

Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Private LastItemSelected As Integer

Private Sub InternalList_Click()
    If LastItemSelected <> InternalList.ListIndex Then
        LastItemSelected = InternalList.ListIndex
        RaiseEvent SelectionChanged
    End If
End Sub

Private Sub InternalList_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub InternalList_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub InternalList_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_EnterFocus()
    InternalList.SetFocus
End Sub

Private Sub UserControl_Initialize()
    InternalList.Left = 0
    InternalList.Top = 0
    InternalList.Width = ScaleWidth
    InternalList.Height = ScaleHeight
    
    CooldownListCount = 0
    
    DefaultCoolDown = 1
    
    BarColor = &HFF
    BarBackColor = &HFFFFFF
    
    Brush = CreateSolidBrush(BarColor)
    BackBrush = CreateSolidBrush(BarBackColor)
    
End Sub
Public Property Get BarOffsetY() As Integer
  BarOffsetY = mBarOffsetY
End Property

Public Property Let BarOffsetY(ByVal NewBarOffsetY As Integer)
  mBarOffsetY = NewBarOffsetY
  PropertyChanged "BarOffsetY"
End Property

Public Property Get BarColor() As OLE_COLOR
  BarColor = mBarColor
End Property

Public Property Let BarColor(ByVal NewBarColor As OLE_COLOR)
  mBarColor = NewBarColor
  PropertyChanged "BarColor"
  Call UpdateBrushes
End Property

Public Property Get BarBackColor() As OLE_COLOR
    BarBackColor = mBarBackColor
End Property

Public Property Let BarBackColor(ByVal NewBarBackColor As OLE_COLOR)
    mBarBackColor = NewBarBackColor
    PropertyChanged "BarBackColor"
    Call UpdateBrushes
End Property

Public Property Get ListIndex() As Integer
    ListIndex = InternalList.ListIndex
End Property
Public Property Let ListIndex(ByVal Index As Integer)
    InternalList.ListIndex = Index
End Property
Public Property Get ListCount() As Integer
    ListCount = InternalList.ListCount
End Property

Public Property Get List(ByVal Index As Integer) As String
    List = InternalList.List(Index)
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = InternalList.BackColor
End Property
Public Property Let BackColor(ByVal NewColor As OLE_COLOR)
    InternalList.BackColor = NewColor
    UserControl.BackColor = NewColor
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = InternalList.ForeColor
End Property
Public Property Let ForeColor(ByVal NewColor As OLE_COLOR)
    InternalList.ForeColor = NewColor
End Property

Public Property Get Font() As StdFont
   Set Font = InternalList.Font
End Property

Public Property Set Font(newFont As StdFont)
   Set UserControl.Font = newFont
   Set InternalList.Font = newFont

   PropertyChanged "Font"
   
   InternalList.Refresh
End Property

Private Sub UserControl_InitProperties()
   Set InternalList.Font = UserControl.Font
End Sub

Public Sub Clean()
    Dim I As Integer
    InternalList.Clear
    
    
    For I = 0 To UBound(CooldownList)
        With CooldownList(I)
            .Cooldown = 0
            .Name = ""
            .startTime = 0
        End With
    Next I
    
    CooldownListCount = 0
    DefaultCooldownStartTime = 0
End Sub


Private Sub UpdateBrushes()
    If Brush > 0 Then
        DeleteObject (Brush)
    End If
    If BackBrush > 0 Then
        DeleteObject (BackBrush)
    End If

    Brush = CreateSolidBrush(BarColor)
    BackBrush = CreateSolidBrush(BarBackColor)
    
End Sub
Private Function GetHeightOfListItem(pListbox As ListBox, pListItem As Integer) As Long
    If pListbox.ListCount = 0 Then Exit Function
    GetHeightOfListItem = SendMessage(pListbox.hwnd, LB_GETITEMHEIGHT, pListItem, 0)
End Function

Private Sub TimerRender_Timer()
    Call Render
End Sub

Public Sub Render()
    If InternalList.ListCount = 0 Then Exit Sub
    Dim RECT As RECT
    Dim Index As Long
    Dim Tick As Long
    Dim Elapsed As Long
    Dim Percent As Double
    Dim StepY As Integer
    Dim Y As Integer
    Dim Z As Long
    Dim CooldownTouse As Long
    StepY = ItemHeight
    
    Tick = GetTickCount
    For Index = InternalList.TopIndex To InternalList.TopIndex + (InternalList.Height / ItemHeight)
        With CooldownList(Index)
            If .startTime > 0 Then
                CooldownTouse = IIf(.ShowOnlyDefault, DefaultCoolDown, .Cooldown)
                Elapsed = CooldownTouse - (Tick - .startTime)
                Percent = ((100 * Elapsed) / CooldownTouse) / 100

                RECT.Left = 0
                RECT.Top = Y + ItemHeight - 2
                RECT.Bottom = RECT.Top + 2
                
                RECT.Right = (InternalList.Width - 16)
                FillRect ListHdc, RECT, BackBrush
                
                RECT.Right = (InternalList.Width - 16) * Percent
                FillRect ListHdc, RECT, Brush
                
                If Percent <= 0 Then
                    .startTime = 0
                    Call InternalList.Refresh
                    RaiseEvent CooldownFinish(CInt(Index))
                End If
            End If
            
        End With
        
        Y = Y + StepY
    Next Index
    
    ' Now calculate the default, if it's set.
    If DefaultCooldownStartTime > 0 Then
    
        Elapsed = DefaultCoolDown - (Tick - DefaultCooldownStartTime)
        Percent = ((100 * Elapsed) / DefaultCoolDown) / 100
    
        If Percent <= 0 Then
            DefaultCooldownStartTime = 0
            RaiseEvent CooldownFinish(-1)
        End If
    
    End If
    
    
    If SpellAfterMeleeCoolDownStartTime > 0 Then
        Elapsed = SpellAfterMeleeCoolDown - (Tick - SpellAfterMeleeCoolDownStartTime)
        Percent = ((100 * Elapsed) / SpellAfterMeleeCoolDown) / 100
    
        If Percent <= 0 Then
            SpellAfterMeleeCoolDownStartTime = 0
            RaiseEvent SpellAfterMeleeCooldownFinish
        End If
    
    End If

    
    
End Sub

Public Property Get ItemIsReady(ByVal ItemIndex As Integer) As Integer
    If ItemIndex < 0 Then
        ItemIsReady = False
        Exit Property
    End If
        
    ItemIsReady = CooldownList(ItemIndex).startTime <= 0
End Property

Public Property Get SpellAfterMeleeIsReady() As Integer
    SpellAfterMeleeIsReady = SpellAfterMeleeCoolDownStartTime <= 0
End Property

Public Property Get DefaultIsReady() As Integer
    DefaultIsReady = DefaultCooldownStartTime <= 0
End Property

Public Property Get SelectedIsReady() As Integer
    If InternalList.ListIndex < 0 Then
        SelectedIsReady = False
        Exit Property
    End If
    
    SelectedIsReady = CooldownList(InternalList.ListIndex).startTime <= 0
End Property

Public Property Get GetName(ByVal ItemIndex As Integer) As String
    If ItemIndex = -1 Then
        GetName = "DefaultValue"
        Exit Property
    End If
    
    GetName = CooldownList(ItemIndex).Name
End Property

Public Sub SetSpellAfterMeleeCooldown(ByVal CoolDownInMs As Long)
    SpellAfterMeleeCoolDown = CoolDownInMs
End Sub

Public Sub SetDefaultCooldown(ByVal CoolDownInMs As Integer)
    DefaultCoolDown = CoolDownInMs
End Sub

Public Sub Start(ByVal Index As Integer, Optional ByVal OnlyDefault As Boolean = False)
    If Index < 0 Then
        Exit Sub
    End If
    
    ' We also need to set the "default timeout", no matter if we have a cooldown defined on the spell we're casting
    DefaultCooldownStartTime = GetTickCount
    
    If CooldownList(Index).Cooldown <= 0 Then Exit Sub
    
    CooldownList(Index).startTime = DefaultCooldownStartTime
    CooldownList(Index).ShowOnlyDefault = OnlyDefault
    
End Sub

Public Sub StartSpellAfterMelee()
    ' We also need to set the "default timeout", no matter if we have a cooldown defined on the spell we're casting
    SpellAfterMeleeCoolDownStartTime = GetTickCount
    
End Sub

Public Sub SetItem(ByVal Index As Integer, ByRef Name As String, ByVal CooldownMs As Long)
    With CooldownList(Index - 1)
        .Name = Name
        .Cooldown = CooldownMs
        .startTime = 0
    End With
    
    InternalList.List(Index - 1) = Name
    
    If ListHdc = 0 Then
        ItemHeight = GetHeightOfListItem(InternalList, 0)
        ListHdc = GetDC(InternalList.hwnd)
        TimerRender.Enabled = True
    End If
    
End Sub
Public Sub Add(ByRef Name As String, ByVal CooldownMs As Long)
    With CooldownList(CooldownListCount)
        .Name = Name
        .Cooldown = CooldownMs
        .startTime = 0
    End With
    Call InternalList.AddItem(Name)
    CooldownListCount = CooldownListCount + 1
    If ListHdc = 0 Then
        ItemHeight = GetHeightOfListItem(InternalList, 0)
        ListHdc = GetDC(InternalList.hwnd)
        TimerRender.Enabled = True
    End If
End Sub

Private Sub UserControl_Resize()
    If ScaleWidth <= 0 Then
        Exit Sub
    End If
    InternalList.Left = 0
    InternalList.Top = 0
    InternalList.Width = ScaleWidth '- 15
    InternalList.Height = ScaleHeight
End Sub

Public Function TryMoveItemUp() As Boolean
    If InternalList.ListIndex < 0 Then
        TryMoveItemUp = False
        Exit Function
    End If
    
    Dim NewIndex As Integer
    Dim OldIndex As Integer
    
    NewIndex = InternalList.ListIndex - 1
    OldIndex = InternalList.ListIndex
    
    Dim TempItem As CooldownListData
    TempItem = CooldownList(NewIndex)
    CooldownList(NewIndex) = CooldownList(OldIndex)
    CooldownList(OldIndex) = TempItem
    
    InternalList.List(NewIndex) = CooldownList(NewIndex).Name
    InternalList.List(OldIndex) = CooldownList(OldIndex).Name
    InternalList.ListIndex = NewIndex
    TryMoveItemUp = True
End Function

Public Function TryMoveItemDown() As Boolean
    If InternalList.ListIndex + 1 >= InternalList.ListCount Then
        TryMoveItemDown = False
        Exit Function
    End If
    
    Dim NewIndex As Integer
    Dim OldIndex As Integer
    
    NewIndex = InternalList.ListIndex + 1
    OldIndex = InternalList.ListIndex
    
    Dim TempItem As CooldownListData
    TempItem = CooldownList(NewIndex)
    CooldownList(NewIndex) = CooldownList(OldIndex)
    CooldownList(OldIndex) = TempItem
    
    InternalList.List(NewIndex) = CooldownList(NewIndex).Name
    InternalList.List(OldIndex) = CooldownList(OldIndex).Name
    InternalList.ListIndex = NewIndex
    TryMoveItemDown = True
End Function

Private Sub UserControl_Terminate()
    Call DeleteObject(Brush)
    Call DeleteObject(BackBrush)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mBarColor = PropBag.ReadProperty("BarColor", &HFF)
    mBarBackColor = PropBag.ReadProperty("BarBackColor", &HFFFFFFFF)
    
    mBarOffsetY = PropBag.ReadProperty("BarOffsetY", 0)
    
    InternalList.BackColor = PropBag.ReadProperty("BackColor", vbWhite)
    InternalList.ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
    
    UserControl.BackColor = InternalList.BackColor
    
    Dim o As Variant
    
    Set o = PropBag.ReadProperty("Font", Nothing)
    If Not o Is Nothing Then
        Set InternalList.Font = o
    End If
    
    Call UpdateBrushes
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Call PropBag.WriteProperty("BarColor", mBarColor, &HFF)
    Call PropBag.WriteProperty("BarBackColor", mBarBackColor, &HFFFFFFFF)
    Call PropBag.WriteProperty("BackColor", InternalList.BackColor, vbWhite)
    Call PropBag.WriteProperty("ForeColor", InternalList.ForeColor, vbBlack)
    Call PropBag.WriteProperty("BarOffsetY", BarOffsetY)
    Call PropBag.WriteProperty("Font", InternalList.Font)

End Sub
