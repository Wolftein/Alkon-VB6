VERSION 5.00
Begin VB.Form frmGuildMembers 
   BackColor       =   &H00292929&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7815
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleMode       =   0  'User
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll 
      Height          =   4815
      Left            =   6750
      Max             =   100
      TabIndex        =   2
      Top             =   120
      Width           =   250
   End
   Begin VB.PictureBox PboContainer 
      BackColor       =   &H00292929&
      BorderStyle     =   0  'None
      Height          =   4850
      Left            =   960
      ScaleHeight     =   4845
      ScaleWidth      =   5685
      TabIndex        =   0
      Top             =   120
      Width           =   5680
      Begin VB.PictureBox PboMembers 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00292929&
         BorderStyle     =   0  'None
         Height          =   6000
         Left            =   -600
         ScaleHeight     =   6000
         ScaleWidth      =   6555
         TabIndex        =   1
         Top             =   0
         Width           =   6550
      End
   End
   Begin VB.Label lblGuildMemberQty 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7080
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.Image ImgInvite 
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   7080
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "frmGuildMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const ControlNamePattern As String = "roleGroup_"
Private CreatedControls() As String

Private cButtonInvite As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()

    Dim I As Integer
    Dim LastControlPosition As Long

    Call LoadControls
    Call InviteButtonEnable
    
    Call modCustomCursors.SetFormCursorDefault(Me)
End Sub
Public Sub InviteButtonEnable()
    Dim AddMemberPermission As Boolean
    
    AddMemberPermission = HasPermission(GP_MEMBER_ACCEPT)
    ImgInvite.Enabled = AddMemberPermission
    cButtonInvite.EnableButton (AddMemberPermission)
    
    Exit Sub
End Sub
Public Sub InitializeUCs()
    If ((Not PlayerData.Guild.Members) = -1) Then
        Exit Sub
    End If
    
    LastControlPosition = 0
    
    ReDim CreatedControls(1 To UBound(PlayerData.Guild.Roles))
    
    For I = 1 To UBound(PlayerData.Guild.Roles)
        CreatedControls(I) = ControlNamePattern & PlayerData.Guild.Roles(I).RoleId
        LastControlPosition = LastControlPosition + ShowRole(PlayerData.Guild.Roles(I), LastControlPosition) + 80
    Next I
    
    If frmGuildMembers.PboContainer.Height > LastControlPosition Then
        frmGuildMembers.VScroll.Visible = False
    Else
        frmGuildMembers.VScroll.Visible = True
    End If
    
    Dim TempValue As Integer
    TempValue = frmGuildMain.PboContainer.Height / LastControlPosition * 100
    
    If TempValue <= 0 Then
        TempValue = 80
    End If
    
    VScroll.LargeChange = TempValue
    PboMembers.Height = LastControlPosition
    PboMembers.Left = 0
    PboMembers.Top = 0
End Sub

Public Sub CleanUCs()
    If ((Not CreatedControls) = -1) Then
        Exit Sub
    End If
    
    Dim I As Integer
    For I = 1 To UBound(CreatedControls)
        Call Me.Controls.Remove(CreatedControls(I))
    Next I
End Sub

Private Function ShowRole(ByRef Role As tGuildRole, ByVal Position As Long) As Long
    Dim MemberListCtl As Control
    Dim RolesMembersCount, MemberIndex As Integer
    Set MemberListCtl = Controls.Add("ARGENTUM.ucGuildRole", ControlNamePattern & Role.RoleId)
    
    With MemberListCtl
        
        RolesMembersCount = 0
        MemberIndex = 0
        ' Fix this. If the UserControl full name is bigger than 39 chars, this won't work
        ' because of a runtime error 1741: https://windows10dll.nirsoft.net/msvbvm60_dll.html
        .Top = IIf(LastControlPosition = 1, 0, LastControlPosition + .Height - 15)
        .Left = 0
        .Visible = True
        
        SetParent .hwnd, PboMembers.hwnd
        .Move 0, Position, .Width, .Height

        For I = 1 To UBound(PlayerData.Guild.Members)
            If PlayerData.Guild.Members(I).RoleId = Role.RoleId Then
                RolesMembersCount = RolesMembersCount + 1
            End If
        Next I
        
        Call MemberListCtl.SetRole(Role.RoleName, RolesMembersCount)
        
        If RolesMembersCount > 0 Then
            For I = 1 To UBound(PlayerData.Guild.Members)
                With PlayerData.Guild.Members(I)
                    If .RoleId = Role.RoleId Then
                        MemberIndex = MemberIndex + 1
                        Call MemberListCtl.AddMember(MemberIndex, .UserId, .UserName, .IsOnline, .RoleId)
                    End If
                End With
            Next I
            
            
        End If
        Call MemberListCtl.CreateMembersControls
        
    End With
    
    ShowRole = MemberListCtl.Height
    
    DoEvents
End Function

Private Sub DoScroll()
    Dim Top As Double
    If (PboMembers.Height - PboContainer.Height) > 0 Then
        Top = (PboMembers.Height - PboContainer.Height) * VScroll.value / 100
        PboMembers.Top = -Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set PboMembers = Nothing
    Set MemberListCtl = Nothing
    Erase CreatedControls
End Sub

Private Sub ImgInvite_Click()
    Call frmGuildMain.LoadForm(frmGuildMemberAdd, "Agregar Miembro")
End Sub

Private Sub VScroll_Change()
    Call DoScroll
End Sub

Private Sub VScroll_Scroll()
    Call DoScroll
End Sub

Public Sub MemberListUpdate(ByVal IdUser As Long, ByVal OnlineStatus As Boolean)
    
    Dim controlesForm As Control

    Call UpdateMemberQty

    For Each controlesForm In Controls
      If (TypeOf controlesForm Is ucGuildRole) Then
           Call controlesForm.ChangeMember(IdUser, OnlineStatus)
      End If
    Next
    
    Exit Sub
End Sub

Public Sub LoadControls()
    Set cButtonInvite = New clsGraphicalButton

    Set LastButtonPressed = New clsGraphicalButton
    
    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI
    
    Call UpdateMemberQty
    
    Call cButtonInvite.Initialize(ImgInvite, GrhPath & "BotonGuildMembersInvite.jpg", _
                                    GrhPath & "BotonGuildMembersInvite.jpg", _
                                    GrhPath & "BotonGuildMembersInvite.jpg", Me)
    '                                'GrhPath & "BotonGuildMembersInvite_Rollover.jpg", _
    '                                'GrhPath & "BotonGuildMembersInvite_Click.jpg", Me)
    '

    Call ImgInvite.ZOrder(1)
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CloseWindow
End Sub

Private Sub CloseWindow()
On Error GoTo ErrHandler
  
    Unload Me
    If frmGuildMain.Visible Then
        Unload frmGuildMain
    End If
    If frmMain.Visible Then frmMain.SetFocus
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmGuildMembers.frm")
End Sub

Public Sub UpdateMemberQty()
    lblGuildMemberQty = PlayerData.Guild.MemberCount & "/" & PlayerData.Guild.MaxMemberQty
End Sub
