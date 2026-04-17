VERSION 5.00
Begin VB.Form frmGuildRolesPermissions 
   BackColor       =   &H00292929&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   316.29
   ScaleMode       =   0  'User
   ScaleWidth      =   543
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRoleName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00292929&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Text            =   "Nombre del Rol"
      Top             =   360
      Width           =   4575
   End
   Begin VB.Image imgButtonSave 
      Height          =   525
      Left            =   3480
      Top             =   4200
      Width           =   1230
   End
   Begin VB.Image imgCheckMemberKick 
      Height          =   210
      Left            =   5595
      Top             =   3428
      Width           =   210
   End
   Begin VB.Image imgCheckMemberInvite 
      Height          =   210
      Left            =   5595
      Top             =   3089
      Width           =   210
   End
   Begin VB.Image imgCheckBankGoldWithdrawn 
      Height          =   210
      Left            =   3330
      Top             =   3428
      Width           =   210
   End
   Begin VB.Image imgCheckBankGoldDeposit 
      Height          =   210
      Left            =   3330
      Top             =   3089
      Width           =   210
   End
   Begin VB.Image imgCheckBankItemWithdrawn 
      Height          =   210
      Left            =   1140
      Top             =   3428
      Width           =   210
   End
   Begin VB.Image imgCheckBankItemDeposit 
      Height          =   210
      Left            =   1140
      Top             =   3089
      Width           =   210
   End
   Begin VB.Image imgCheckRoleEdit 
      Height          =   210
      Left            =   5595
      Top             =   2130
      Width           =   210
   End
   Begin VB.Image imgCheckRoleCreateDelete 
      Height          =   210
      Left            =   5595
      Top             =   1800
      Width           =   210
   End
   Begin VB.Image imgAssignRoleOthers 
      Height          =   210
      Left            =   3345
      Top             =   2130
      Width           =   210
   End
   Begin VB.Image imgAssignRoleRightHand 
      Height          =   210
      Left            =   3345
      Top             =   1800
      Width           =   210
   End
   Begin VB.Image imgCheckEditDesc 
      Height          =   210
      Left            =   1095
      Top             =   1800
      Width           =   210
   End
End
Attribute VB_Name = "frmGuildRolesPermissions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cButtonCheckEditDesc As clsGraphicalButton
Private cButtonAssignRoleRightHand As clsGraphicalButton
Private cButtonAssignRoleOthers As clsGraphicalButton
Private cButtonRoleCreateDelete As clsGraphicalButton
Private cButtonRoleEdit As clsGraphicalButton
Private cButtonBankItemDeposit As clsGraphicalButton
Private cButtonBankItemWithdrawn As clsGraphicalButton
Private cButtonBankGoldDeposit As clsGraphicalButton
Private cButtonBankGoldWithdrawn As clsGraphicalButton
Private cButtonMemberInvite As clsGraphicalButton
Private cButtonMemberKick As clsGraphicalButton

Private picCheckBox As Picture
Private picCheckBoxDisabled As Picture

Private cButtonSave As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton


Dim QtyPermission As Integer
Dim OldPermissions() As String
Dim NewPermissions() As String
Dim OldRoleName As String
Dim RoleId As Integer

Private CanDeleteRole As Boolean
Private CanEditRolePermission As Boolean
Private CanRenameRole As Boolean

Public Sub LoadControls()
    Set cButtonCheckEditDesc = New clsGraphicalButton
    Set cButtonAssignRoleRightHand = New clsGraphicalButton
    Set cButtonAssignRoleOthers = New clsGraphicalButton
    Set cButtonRoleCreateDelete = New clsGraphicalButton
    Set cButtonRoleEdit = New clsGraphicalButton
    Set cButtonBankItemDeposit = New clsGraphicalButton
    Set cButtonBankItemWithdrawn = New clsGraphicalButton
    Set cButtonBankGoldDeposit = New clsGraphicalButton
    Set cButtonBankGoldWithdrawn = New clsGraphicalButton
    Set cButtonMemberInvite = New clsGraphicalButton
    Set cButtonMemberKick = New clsGraphicalButton
    
    Set cButtonSave = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI
    
    Me.Picture = LoadPicture(GrhPath & "VentanaGuildRolePermission.jpg")
    Set picCheckBox = LoadPicture(GrhPath & "BotonGuildCheckbox.jpg")
    Set picCheckBoxDisabled = LoadPicture(GrhPath & "BotonGuildCheckbox_Disabled.jpg")
  
    Call cButtonSave.Initialize(imgButtonSave, GrhPath & "BotonGuildRolePermissionGuardar.jpg", _
                               GrhPath & "BotonGuildRolePermissionGuardar.jpg", _
                               GrhPath & "BotonGuildRolePermissionGuardar.jpg", Me, _
                               GrhPath & "BotonGuildRolePermissionGuardar_Disabled.jpg")

End Sub

Public Sub CleanData()
    QtyPermission = 0
    RoleId = 0
    Erase OldPermissions
    Erase NewPermissions
    QtyPermission = 0
    CanEditRolePermission = False
    CanRenameRole = False
    CanDeleteRole = False
    txtRoleName.Enabled = True
    txtRoleName.text = "Nuevo Rol"
    
    Call CleanButtonToggle(imgCheckEditDesc)
    Call CleanButtonToggle(imgAssignRoleRightHand)
    
    Call CleanButtonToggle(imgAssignRoleRightHand)
    Call CleanButtonToggle(imgAssignRoleOthers)
    Call CleanButtonToggle(imgCheckRoleCreateDelete)
    Call CleanButtonToggle(imgCheckRoleEdit)
    Call CleanButtonToggle(imgCheckBankItemDeposit)
    Call CleanButtonToggle(imgCheckBankItemWithdrawn)
    Call CleanButtonToggle(imgCheckBankGoldDeposit)
    Call CleanButtonToggle(imgCheckBankGoldWithdrawn)
    Call CleanButtonToggle(imgCheckMemberInvite)
    Call CleanButtonToggle(imgCheckMemberKick)

End Sub


Public Sub CheckPermission(ByVal Permission As String, ByVal Index As Integer)
    OldPermissions(Index) = Permission
    Dim PicToUse As Picture

    If CanEditRolePermission Then
        Set PicToUse = picCheckBox
    Else
        Set PicToUse = picCheckBoxDisabled
    End If
    
    Select Case Permission
        Case GP_EDIT_GUILD_DESC
            Call TogglePermission(imgCheckEditDesc, PicToUse)
        Case GP_RIGHT_HAND_ASSIGN
            Call TogglePermission(imgAssignRoleRightHand, PicToUse)
        Case GP_ROLE_ASSIGN
            Call TogglePermission(imgAssignRoleOthers, PicToUse)
        Case GP_ROLE_CREATE_DELETE
            Call TogglePermission(imgCheckRoleCreateDelete, PicToUse)
        Case GP_ROLE_MODIFY
            Call TogglePermission(imgCheckRoleEdit, PicToUse)
        Case GP_BANK_DEPOSIT_ITEM
            Call TogglePermission(imgCheckBankItemDeposit, PicToUse)
        Case GP_BANK_WITHDRAW_ITEM
            Call TogglePermission(imgCheckBankItemWithdrawn, PicToUse)
        Case GP_BANK_DEPOSIT_GOLD
            Call TogglePermission(imgCheckBankGoldDeposit, PicToUse)
        Case GP_BANK_WITHDRAW_GOLD
            Call TogglePermission(imgCheckBankGoldWithdrawn, PicToUse)
        Case GP_MEMBER_ACCEPT
            Call TogglePermission(imgCheckMemberInvite, PicToUse)
        Case GP_MEMBER_KICK
            Call TogglePermission(imgCheckMemberKick, PicToUse)
    End Select
End Sub

Public Sub SetRoleData(ByVal RoleName As String, ByVal Id As Integer, ByVal CanEditPermissions As Boolean, ByVal CanRename As Boolean, ByVal CanDelete As Boolean, ByVal PermissionsCount As Integer)
    txtRoleName.text = RoleName
    OldRoleName = RoleName
    RoleId = Id
    
    CanEditRolePermission = CanEditPermissions
    CanRenameRole = CanRename
    CanDeleteRole = CanDelete
    
    If Id <> 0 And PermissionsCount > 0 Then
        'its an update
        ReDim OldPermissions(1 To PermissionsCount)
    End If
       
    txtRoleName.Enabled = CanEditPermissions Or CanRenameRole
        
    imgCheckEditDesc.Picture = Nothing
    imgAssignRoleRightHand.Picture = Nothing
    imgAssignRoleOthers.Picture = Nothing
    imgCheckRoleCreateDelete.Picture = Nothing
    imgCheckRoleEdit.Picture = Nothing
    imgCheckBankItemDeposit.Picture = Nothing
    imgCheckBankItemWithdrawn.Picture = Nothing
    imgCheckBankGoldDeposit.Picture = Nothing
    imgCheckBankGoldWithdrawn.Picture = Nothing
    imgCheckMemberInvite.Picture = Nothing
    imgCheckMemberKick.Picture = Nothing
End Sub


Private Sub imgButtonSave_Click()
    Dim RoleName As String
    
    If RoleId <> 0 Then
        ' If there's not permission to edit the role then we exit
        If Not CanEditRolePermission And Not CanRenameRole Then
            Call frmGuildMain.LoadForm(frmGuildRolesList, "Guild Roles")
            Exit Sub
        End If
    End If
    
    RoleName = Trim(txtRoleName.text)
    
    If RoleName = vbNullString Then
        MsgBox ("Seleccione un nombre para el rol.")
        Exit Sub
    End If
    
    Dim NewPermissions() As String
        
    Dim I As Integer
    Dim J As Integer
    Dim PermissionsChanged As Boolean
    Dim PermissionExists As Boolean

    Call AddPermissionIfChecked(NewPermissions, imgCheckEditDesc, GP_EDIT_GUILD_DESC)
    Call AddPermissionIfChecked(NewPermissions, imgAssignRoleRightHand, GP_RIGHT_HAND_ASSIGN)
    
    Call AddPermissionIfChecked(NewPermissions, imgAssignRoleRightHand, GP_RIGHT_HAND_ASSIGN)
    Call AddPermissionIfChecked(NewPermissions, imgAssignRoleOthers, GP_ROLE_ASSIGN)
    Call AddPermissionIfChecked(NewPermissions, imgCheckRoleCreateDelete, GP_ROLE_CREATE_DELETE)
    Call AddPermissionIfChecked(NewPermissions, imgCheckRoleEdit, GP_ROLE_MODIFY)
    Call AddPermissionIfChecked(NewPermissions, imgCheckBankItemDeposit, GP_BANK_DEPOSIT_ITEM)
    Call AddPermissionIfChecked(NewPermissions, imgCheckBankItemWithdrawn, GP_BANK_WITHDRAW_ITEM)
    Call AddPermissionIfChecked(NewPermissions, imgCheckBankGoldDeposit, GP_BANK_DEPOSIT_GOLD)
    Call AddPermissionIfChecked(NewPermissions, imgCheckBankGoldWithdrawn, GP_BANK_WITHDRAW_GOLD)
    Call AddPermissionIfChecked(NewPermissions, imgCheckMemberInvite, GP_MEMBER_ACCEPT)
    Call AddPermissionIfChecked(NewPermissions, imgCheckMemberKick, GP_MEMBER_KICK)
    
    PermissionsChanged = False
    
    'check if data has changed
    If RoleId <> 0 Then
    
        'check if name changed
        If txtRoleName.text <> OldRoleName Then
            PermissionsChanged = True
        End If
        
        'check if permissions changed
        If Not PermissionsChanged Then
            'before everything, check if there is more or less permissions
            Dim PermissionCount As Integer
            Dim OldPermissionCount As Integer
            If Utility.IsArrayNull(NewPermissions) = False Then PermissionCount = UBound(NewPermissions)
            If Utility.IsArrayNull(OldPermissions) = False Then OldPermissionCount = UBound(OldPermissions)
            
            
            If PermissionCount = OldPermissionCount Then
                'check any permission
                For I = 1 To PermissionCount
                    PermissionExists = False
                    For J = 1 To OldPermissionCount
                        If NewPermissions(I) = OldPermissions(J) Then
                            PermissionExists = True
                        End If
                    Next J
                    If Not PermissionExists Then
                        PermissionsChanged = True
                        Exit For
                    End If
                Next I
            Else
                'if there is more or less permissions, so permission had changed
                PermissionsChanged = True
            End If
        End If
    Else
        'it's a new role, we must save it
        PermissionsChanged = True
    End If
    
    If PermissionsChanged Then
        Call WriteGuildRole_Create(eRoleAction.Create, RoleId, RoleName, NewPermissions)
    End If
    
    Call frmGuildMain.LoadForm(frmGuildRolesList, "Guild Roles")
    
End Sub

Private Sub CheckPermissionsCount(ByRef Cbo As clsGraphicalButton)
    If Cbo.IsEnabled() Then
        QtyPermission = QtyPermission + 1
    ElseIf QtyPermission > 0 Then
        QtyPermission = QtyPermission - 1
    End If
End Sub


Private Sub Form_Load()
    
    Call LoadControls
    
    Call modCustomCursors.SetFormCursorDefault(Me)
End Sub

Private Sub imgCheckBankGoldDeposit_Click()
    If Not CanEditRolePermission Then Exit Sub
    
    Call TogglePermission(imgCheckBankGoldDeposit, picCheckBox)
End Sub

Private Sub imgCheckBankGoldWithdrawn_Click()
    If Not CanEditRolePermission Then Exit Sub
    
    Call TogglePermission(imgCheckBankGoldWithdrawn, picCheckBox)
End Sub

Private Sub imgCheckBankItemDeposit_Click()
    If Not CanEditRolePermission Then Exit Sub
    
    Call TogglePermission(imgCheckBankItemDeposit, picCheckBox)
End Sub

Private Sub imgCheckBankItemWithdrawn_Click()
    If Not CanEditRolePermission Then Exit Sub
    
    Call TogglePermission(imgCheckBankItemWithdrawn, picCheckBox)
End Sub

Private Sub imgCheckEditDesc_Click()
    If Not CanEditRolePermission Then Exit Sub
    
    Call TogglePermission(imgCheckEditDesc, picCheckBox)
End Sub

Private Sub imgAssignRoleRightHand_Click()
    If Not CanEditRolePermission Then Exit Sub
    
    Call TogglePermission(imgAssignRoleRightHand, picCheckBox)
End Sub

Private Sub imgCheckRoleEdit_Click()
    If Not CanEditRolePermission Then Exit Sub
   
     Call TogglePermission(imgCheckRoleEdit, picCheckBox)
End Sub

Private Sub imgAssignRoleOthers_Click()
    If Not CanEditRolePermission Then Exit Sub
    
    Call TogglePermission(imgAssignRoleOthers, picCheckBox)
End Sub

Private Sub imgCheckMemberInvite_Click()
    If Not CanEditRolePermission Then Exit Sub
    
    Call TogglePermission(imgCheckMemberInvite, picCheckBox)
End Sub

Private Sub imgCheckMemberKick_Click()
    If Not CanEditRolePermission Then Exit Sub
    
    Call TogglePermission(imgCheckMemberKick, picCheckBox)
End Sub

Private Sub AddPermissionIfChecked(ByRef Permissions() As String, ByRef Image As Image, ByVal Key As String)
    Dim CurrentElementAt As Integer
    
    If Utility.IsArrayNull(Permissions) = False Then CurrentElementAt = UBound(Permissions)
    
    If IsCheckToggled(Image) Then
        CurrentElementAt = CurrentElementAt + 1
    
        ReDim Preserve Permissions(1 To CurrentElementAt)
        Permissions(CurrentElementAt) = Key
    End If
End Sub

Private Sub imgCheckRoleCreateDelete_Click()
    If Not CanEditRolePermission Then Exit Sub
    
    Call TogglePermission(imgCheckRoleCreateDelete, picCheckBox)
End Sub

Private Sub CleanButtonToggle(ByRef Image As Image)
    Image.Tag = ""
    Image.Picture = Nothing
End Sub

Private Sub TogglePermission(ByRef Image As Image, ByRef ImageToUse As Picture)
    Dim Enabled As Boolean

    Enabled = IsCheckToggled(Image)
    
    If Enabled Then
        Image.Tag = "0"
        Image.Picture = Nothing
    Else
        Image.Tag = "1"
        Image.Picture = ImageToUse
    End If
      
End Sub

Private Function IsCheckToggled(ByRef Image As Image) As Boolean
    
    IsCheckToggled = CBool(IIf(Image.Tag = "" Or Image.Tag = "0", False, True))

End Function

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
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmGuildRolesPermissions.frm")
End Sub
