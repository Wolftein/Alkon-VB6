VERSION 5.00
Begin VB.UserControl ucGuildUserMember 
   BackColor       =   &H000000FF&
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5685
   ScaleHeight     =   525
   ScaleMode       =   0  'User
   ScaleWidth      =   5541.514
   Begin VB.Image ImgKickMember 
      Height          =   450
      Left            =   4606
      Top             =   60
      Width           =   450
   End
   Begin VB.Label LblUsername 
      BackStyle       =   0  'Transparent
      Caption         =   "Nick miembro"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   3975
   End
   Begin VB.Image ImgRoleAssign 
      Height          =   450
      Left            =   4104
      Top             =   60
      Width           =   495
   End
   Begin VB.Shape ShpIsOnline 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   360
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   80
      Width           =   605
   End
End
Attribute VB_Name = "ucGuildUserMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public UserName As String
Public IdRole As Integer
Public IdUser As Long
Private cButtonKick As clsGraphicalButton
Private cButtonAssignRole As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton


Public Sub LoadData(ByVal UserName As String, ByVal RoleId As Integer, ByVal UserId As Long, ByVal IsOnline As Boolean, ByVal CanEdit As Boolean)
    LblUsername.Alignment = AlignmentConstants.vbLeftJustify
    LblUsername.Caption = UserName
    
    ShpIsOnline.Visible = True
    ShpIsOnline.BackColor = IIf(IsOnline, &HFF00&, &HFF&)
    
    ImgKickMember.Visible = CanEdit
    ImgKickMember.Enabled = CanEdit
    Call cButtonKick.EnableButton(CanEdit)
    
    Dim CanAssignRole As Boolean
    CanAssignRole = HasPermission(GP_ROLE_ASSIGN) Or HasPermission(GP_RIGHT_HAND_ASSIGN)
    
    ImgRoleAssign.Visible = CanAssignRole
    Call cButtonAssignRole.EnableButton(CanAssignRole)
        
    IdRole = RoleId
    IdUser = UserId
    
    If IdRole = ID_ROLE_LEADER Then
        Call cButtonAssignRole.EnableButton(False)
        'Call cButtonKick.EnableButton(False)
    End If
End Sub

Public Sub SetEmpty()
    LblUsername.Alignment = AlignmentConstants.vbCenter
    LblUsername.Caption = "-"
    ShpIsOnline.Visible = False
    ImgKickMember.Visible = False
    ImgRoleAssign.Visible = False
End Sub

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Private Sub ImgKickMember_Click()
    If Not cButtonKick.IsEnabled Then
        Exit Sub
    End If
    
    Call KickGuildMember(LblUsername.Caption)
End Sub

Private Sub ImgRoleAssign_Click()
    If Not cButtonAssignRole.IsEnabled Then
        Exit Sub
    End If
    Call frmGuildMain.LoadForm(frmGuildRoleAssign, "Asignación de roles")
    Call frmGuildRoleAssign.LoadRoles(IdRole, IdUser)
End Sub

Private Sub UserControl_Initialize()
        
    Set cButtonKick = New clsGraphicalButton
    Set cButtonAssignRole = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    
    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI
        
    Call cButtonKick.Initialize(ImgKickMember, GrhPath & "BotonGuildMemberKick.jpg", _
                                    GrhPath & "BotonGuildMemberKick.jpg", _
                                    GrhPath & "BotonGuildMemberKick.jpg", frmGuildMembers)
                                    'GrhPath & "BotonGuildMemberKick.jpg", _
                                    'GrhPath & "BotonGuildMemberKick.jpg", Me)
                                    
    'TODO image
    Call cButtonAssignRole.Initialize(ImgRoleAssign, GrhPath & "BotonGuildMemberConfig.jpg", _
                                    GrhPath & "BotonGuildMemberConfig.jpg", _
                                    GrhPath & "BotonGuildMemberConfig.jpg", frmGuildMembers)
                                    'GrhPath & "BotonGuildMemberConfig.jpg", _
                                    'GrhPath & "BotonGuildMemberConfig.jpg", Me)
                            
        
    UserControl.Picture = LoadPicture(GrhPath & "VentanaGuildBankMemberItem.jpg")
    
End Sub



