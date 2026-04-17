VERSION 5.00
Begin VB.UserControl ucGuildRole 
   BackColor       =   &H00292929&
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5685
   MaskColor       =   &H00292929&
   ScaleHeight     =   495
   ScaleWidth      =   5685
   Begin VB.Label lblRoleName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Un rol"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   50
      Width           =   5415
   End
End
Attribute VB_Name = "ucGuildRole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Members() As tGuildUserMember
Private MemberCount As Integer

Private Const ControlNamePattern As String = "memberUserCtrl_"
Private Const FirstElementPosition As Integer = 680

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Sub SetRole(ByVal RoleName As String, ByVal RoleMemberCount As Integer)
    lblRoleName.Caption = RoleName
    If RoleMemberCount <> 0 Then
        ReDim Members(1 To RoleMemberCount)
    End If
    
    MemberCount = RoleMemberCount
End Sub

Public Sub AddMember(ByVal Index As Integer, ByVal UserId As Long, ByVal UserName As String, ByVal IsOnline As Boolean, ByVal RoleId As Integer)
    With Members(Index)
        .IsOnline = IsOnline
        .UserId = UserId
        .UserName = UserName
        .RoleId = RoleId
    End With
       
End Sub

Public Sub CreateMembersControls()

    Dim Obj As Control
    Dim myCtl As Control

    Dim I As Integer
    Dim KickMemberPermission As Boolean
    Dim KickButtonVisible As Boolean
    
    Call RemoveControls
    
    Dim ControlHeight As Long
    Dim LastControlPosition As Long
    Dim HasEditPermission As Integer
    LastControlPosition = 0
    KickMemberPermission = HasPermission(GP_MEMBER_KICK)
    
    If MemberCount = 0 Then
        ControlHeight = CreateEmptyControl()
        UserControl.Height = FirstElementPosition + ControlHeight
        Exit Sub
    End If
    
    For I = 1 To UBound(Members)
        Set myCtl = Controls.Add("ARGENTUM.ucGuildUserMember", ControlNamePattern & I)
        
        ControlHeight = myCtl.Height
        
        myCtl.Top = IIf(LastControlPosition = 0, FirstElementPosition, LastControlPosition + ControlHeight - 15)
        myCtl.Left = 0
        myCtl.Visible = True
        
        SetParent myCtl.hwnd, Me.hwnd
        myCtl.Move myCtl.Left, myCtl.Top, myCtl.Width, myCtl.Height

        KickButtonVisible = KickMemberPermission Or (Members(I).UserName = UserName)
        
        If Members(I).RoleId = ID_ROLE_LEADER And PlayerData.Guild.IdRolOwn <> ID_ROLE_LEADER Then KickButtonVisible = False

        Call myCtl.LoadData(Members(I).UserName, Members(I).RoleId, Members(I).UserId, Members(I).IsOnline, KickButtonVisible)
        
        LastControlPosition = myCtl.Top
        
        DoEvents
    Next I
        
     UserControl.Height = FirstElementPosition + (ControlHeight * UBound(Members))
    
    
End Sub

Private Function CreateEmptyControl() As Integer
    Dim myCtl As Control
    Set myCtl = Controls.Add("ARGENTUM.ucGuildUserMember", ControlNamePattern & "1")
    
    myCtl.Top = IIf(LastControlPosition = 0, FirstElementPosition, myCtl.Height + ControlHeight - 15)
    myCtl.Left = 0
    myCtl.Visible = True
    
    SetParent myCtl.hwnd, Me.hwnd
    myCtl.Move myCtl.Left, myCtl.Top, myCtl.Width, myCtl.Height
    
    Call myCtl.SetEmpty
    
    CreateEmptyControl = myCtl.Height

End Function

Private Sub RemoveControls()
    Dim I As Integer
    Dim Control As Control
    For Each Control In Controls
        If InStr(Control.Name, ControlNamePattern) > 0 Then
            Call Controls.Remove(Control.Name)
        End If
    Next
End Sub

Private Sub UserControl_Terminate()
    Set myCtl = Nothing
End Sub

Public Sub ChangeMember(ByVal IdUser As Long, ByVal OnlineStatus As Boolean)
    Dim I As Integer
    
    If ((Not Members) <> -1) Then
        For I = 1 To UBound(Members)
            If Members(I).UserId = IdUser Then
                Members(I).IsOnline = OnlineStatus
                Call RemoveControls
                Call CreateMembersControls
            End If
        Next I
    End If

    Exit Sub
End Sub
