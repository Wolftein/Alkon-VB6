VERSION 5.00
Begin VB.UserControl ucGuildRolePermissions 
   BackColor       =   &H000000FF&
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5685
   ScaleHeight     =   525
   ScaleWidth      =   5685
   Begin VB.Image ImgDeleteRole 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   4740
      Top             =   45
      Width           =   435
   End
   Begin VB.Image ImgEdRole 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   5200
      Top             =   45
      Width           =   435
   End
   Begin VB.Label LblRoleName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "ucGuildRolePermissions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Permissions() As String
Private PermissionsQty As Integer
Private RoleId As Integer

Private CanDeleteRole
Private CanUpdateRolePermissions As Boolean
Private CanRenameRole As Boolean

Private cButtonEditRole As clsGraphicalButton
Private cButtonDeleteRole As clsGraphicalButton
Public LastButtonPressed As clsGraphicalButton

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property



Public Sub ShowRole(ByVal Id As Integer, ByVal Name As String, ByVal PermissionCount As Integer, ByVal CanDelete As Boolean, ByVal CanUpdate As Boolean, ByVal CanRename As Boolean)
    LblRoleName.Caption = Name
    RoleId = Id
    CanEdit = IsDeleteable
    PermissionsQty = PermissionCount
    
    CanDeleteRole = CanDelete
    CanUpdateRolePermissions = CanUpdate
    CanRenameRole = CanRename
    
    
    Call cButtonDeleteRole.EnableButton(CanDeleteRole)
    ImgDeleteRole.Visible = CanDeleteRole
    
    If PermissionCount <= 0 Then Exit Sub
       
    ReDim Permissions(1 To PermissionCount)
End Sub

Public Sub AddPermission(ByVal Index As Integer, ByVal Permission As String)
    Permissions(Index) = Permission
End Sub


Private Sub ImgDeleteRole_Click()
    Call Delete
End Sub

Private Sub ImgEdRole_Click()
    Call Edit
End Sub

Private Sub LblRoleName_Click()
    Call Edit
End Sub

Private Sub UserControl_Click()
    Call Edit
End Sub

Private Sub LoadControls()

    Set cButtonEditRole = New clsGraphicalButton
    Set cButtonDeleteRole = New clsGraphicalButton
    Set LastButtonPressed = New clsGraphicalButton
    
    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI
    
    UserControl.Picture = LoadPicture(GrhPath & "VentanaGuildBankMemberItem.jpg")
    
    
    Call cButtonEditRole.Initialize(ImgEdRole, GrhPath & "BotonGuildMemberConfig.jpg", _
                                    GrhPath & "BotonGuildMemberConfig.jpg", _
                                    GrhPath & "BotonGuildMemberConfig.jpg", frmGuildRolesList, _
                                    GrhPath & "BotonGuildMemberConfig.jpg")
                
    Call cButtonDeleteRole.Initialize(ImgDeleteRole, GrhPath & "BotonGuildRoleRemove.jpg", _
                                    GrhPath & "BotonGuildRoleRemove.jpg", _
                                    GrhPath & "BotonGuildRoleRemove.jpg", frmGuildRolesList, _
                                    GrhPath & "BotonGuildRoleRemove.jpg")
                                    

End Sub

Private Sub Edit()
    Call frmGuildMain.LoadForm(frmGuildRolesPermissions, "Permisos")
    
    Call frmGuildRolesPermissions.CleanData
   
    Call frmGuildRolesPermissions.SetRoleData(LblRoleName.Caption, RoleId, CanUpdateRolePermissions, CanRenameRole, CanDeleteRole, PermissionsQty)
   
    For J = 1 To PermissionsQty
         Call frmGuildRolesPermissions.CheckPermission(Permissions(J), J)
    Next
End Sub

Private Sub Delete()

    If Not HasPermission(GP_ROLE_CREATE_DELETE) Then Exit Sub
    
    Call Protocol.WriteGuildRole_Delete(RoleId)
        
End Sub

Private Sub UserControl_Initialize()
    Call LoadControls
End Sub
