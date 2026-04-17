VERSION 5.00
Begin VB.Form frmGuildRoleAssign 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CboRoles 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   240
      Width           =   3375
   End
   Begin VB.Image ImgSave 
      Height          =   495
      Left            =   4320
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmGuildRoleAssign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cButtonSave As clsGraphicalButton
Private OldRoleId As Integer
Private TargetUserId As Long

Public Sub LoadRoles(ByVal RoleId As Integer, ByVal TargetUser As Long)
    Dim I As Integer
    Call CboRoles.Clear
    OldRoleId = RoleId
    TargetUserId = TargetUser
    For I = 1 To UBound(PlayerData.Guild.Roles)
        If PlayerData.Guild.Roles(I).RoleId = RoleId Then
            'siempre muestro el rol actual como opcion
            Call CboRoles.AddItem(PlayerData.Guild.Roles(I).RoleName)
        Else
            If PlayerData.Guild.Roles(I).RoleId <> ID_ROLE_LEADER Then
                If PlayerData.Guild.Roles(I).RoleId <> ID_ROLE_RIGHTHAND Or PlayerData.Guild.IdRightHand = 0 Then
                    Call CboRoles.AddItem(PlayerData.Guild.Roles(I).RoleName)
                End If
            End If
        End If
       
        
        If PlayerData.Guild.Roles(I).RoleId = RoleId Then
            CboRoles.text = PlayerData.Guild.Roles(I).RoleName
        End If
    Next I
    
    
End Sub

Private Sub Form_Load()

    Set cButtonKick = New clsGraphicalButton
    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI
    
    'TODO image
    Call cButtonKick.Initialize(ImgSave, GrhPath & "BotonClanMemberInfoAceptar.jpg", _
                                    GrhPath & "BotonClanMemberInfoAceptar.jpg", _
                                    GrhPath & "BotonClanMemberInfoAceptar.jpg", frmGuildRoleAssign)
                                    'GrhPath & "BotonClanMemberInfoAceptar.jpg", _
                                    'GrhPath & "BotonClanMemberInfoAceptar.jpg", Me)
                                    
    Call modCustomCursors.SetFormCursorDefault(Me)
                                    
End Sub

Private Sub ImgSave_Click()
    Dim NewRoleId As Integer
    
    Dim I As Integer
    
    For I = 1 To UBound(PlayerData.Guild.Roles)
        If PlayerData.Guild.Roles(I).RoleName = CboRoles.text Then
            NewRoleId = PlayerData.Guild.Roles(I).RoleId
            Exit For
        End If
    Next I
    
    If NewRoleId = 0 Then
        Call MsgBox("Seleccione un rol de la lista")
        Exit Sub
    End If
    
    'lider
    If NewRoleId = ID_ROLE_LEADER Then
        Call MsgBox("Este rol no se puede asignar")
        Exit Sub
    End If
    
    'mano derecha
    If NewRoleId = ID_ROLE_RIGHTHAND And Not HasPermission(GP_RIGHT_HAND_ASSIGN) Then
        Call MsgBox("No posee permisos para asignar este rol")
        Exit Sub
    End If
        
    If OldRoleId <> NewRoleId Then
        Call WriteGuildRole_Assign(eRoleAction.Assign, NewRoleId, TargetUserId)
        Call frmGuildMain.LoadForm(frmGuildMembers, "Miembros de clan")
    End If
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
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmGuildRoleAssign.frm")
End Sub
