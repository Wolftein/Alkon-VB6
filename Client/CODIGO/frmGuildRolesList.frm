VERSION 5.00
Begin VB.Form frmGuildRolesList 
   BackColor       =   &H00292929&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1044.014
   ScaleMode       =   0  'User
   ScaleWidth      =   620.73
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PboRoles 
      BackColor       =   &H00292929&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   12968
      ScaleHeight     =   4935
      ScaleWidth      =   6075
      TabIndex        =   2
      Top             =   240
      Width           =   6079
   End
   Begin VB.PictureBox PboRolesContainer 
      BackColor       =   &H00292929&
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   480
      ScaleHeight     =   4455
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   875
      Width           =   5000
   End
   Begin VB.VScrollBar VScroll 
      Height          =   4935
      Left            =   6120
      Max             =   100
      TabIndex        =   0
      Top             =   600
      Width           =   300
   End
   Begin VB.Label lblGuildRolesQty 
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
      Left            =   6480
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
   Begin VB.Image ImgNewRole 
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   6600
      Top             =   600
      Width           =   555
   End
   Begin VB.Label lblFormName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Listado de Roles"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmGuildRolesList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cButtonNewRole As clsGraphicalButton
Public LastButtonPressed As clsGraphicalButton

Private Const ControlNamePattern As String = "roleItem_"
Dim CreatedControls() As String
Private MemberListCtl As Control



Private Sub Form_Load()
    
    If PlayerData.Guild.IdGuild <= 0 Then Exit Sub
    
    Call LoadControls
    
    Call AddRoleButtonEnable
    
    lblGuildRolesQty.Caption = UBound(PlayerData.Guild.Roles) & "/" & PlayerData.Guild.MaxRoles
    
    Call DrawRoles
    Me.Width = 8000
    
    Call modCustomCursors.SetFormCursorDefault(Me)
End Sub

Private Function ShowRole(ByRef Role As tGuildRole, ByVal Position As Long, ByVal Index As Integer) As Long
    Dim MemberListCtl As Control
    Dim ControlName As String
    Dim LastControlPosition As Long
    Dim J As Integer
    
    ControlName = ControlNamePattern & Role.RoleId
    
    If Position = 0 Then Position = 400
    
    Set MemberListCtl = Controls.Add("ARGENTUM.ucGuildRolePermissions", ControlName)
    CreatedControls(Index) = ControlName
    
    With MemberListCtl
    
        ' Fix this. If the UserControl full name is bigger than 39 chars, this won't work
        ' because of a runtime error 1741: https://windows10dll.nirsoft.net/msvbvm60_dll.html
        .Top = IIf(LastControlPosition = 1, 0, LastControlPosition + .Height - 15)
        .Left = 0
        .Visible = True
        
        SetParent .hwnd, PboRoles.hwnd
        .Move 20, Position, .Width, .Height
        
        Call MemberListCtl.ShowRole(Role.RoleId, Role.RoleName, Role.PermissionsQty, Role.DeleteEnabled, Role.UpdatePermissionsEnabled, Role.RenameEnabled)
        
        For J = 1 To Role.PermissionsQty
            Call MemberListCtl.AddPermission(J, Role.Permissions(J).Key)
        Next
    End With
    
    ShowRole = MemberListCtl.Height
    
    DoEvents
End Function

Private Sub DoScroll()
    Dim Top As Double
    Top = (PboRoles.Height - PboRolesContainer.Height) * VScroll.value / 100
    PboRoles.Top = -Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set PboRoles = Nothing
    Set MemberListCtl = Nothing
End Sub

Private Sub ImgNewRole_Click()
    If Not cButtonNewRole.IsEnabled Then Exit Sub
    
    Call frmGuildMain.LoadForm(frmGuildRolesPermissions, "Permisos de Roles")
    
    Call frmGuildRolesPermissions.CleanData
    
    Call frmGuildRolesPermissions.SetRoleData("Nuevo Rol", 0, True, True, False, 11)
End Sub

Private Sub VScroll_Change()
    Call DoScroll
End Sub

Private Sub VScroll_Scroll()
    Call DoScroll
End Sub

Public Sub CleanUCs()
    Dim I As Integer
    For I = 1 To UBound(CreatedControls)
        Me.Controls.Remove (CreatedControls(I))
    Next
End Sub
Public Sub InitializeUCs()
    Call DrawRoles
End Sub

Private Sub LoadControls()

    Set cButtonNewRole = New clsGraphicalButton
    Set LastButtonPressed = New clsGraphicalButton
    
    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI
    
    
    Call cButtonNewRole.Initialize(ImgNewRole, GrhPath & "BotonGuildRoleAdd.jpg", _
                                    GrhPath & "BotonGuildRoleAdd.jpg", _
                                    GrhPath & "BotonGuildRoleAdd.jpg", Me, _
                                    GrhPath & "BotonGuildRoleAdd_Disabled.jpg")
End Sub

Public Sub DrawRoles()
    Dim I As Integer
    Dim LastControlPosition As Long
    Dim ControlSize As Long
    LastControlPosition = 20
    
    
    ReDim CreatedControls(1 To UBound(PlayerData.Guild.Roles)) As String
    
    For I = 1 To UBound(PlayerData.Guild.Roles)
        ControlSize = ShowRole(PlayerData.Guild.Roles(I), LastControlPosition, I)
        LastControlPosition = LastControlPosition + ControlSize
    Next
    
    If frmGuildRolesList.PboRolesContainer.Height > LastControlPosition Then
        VScroll.Visible = False
    End If
    
    VScroll.LargeChange = PboRolesContainer.Height / LastControlPosition * 100
    PboRoles.Height = LastControlPosition
    PboRoles.Left = 0
    PboRoles.Top = 150
    
    'to move scroll to last position in case its a re-load
    Call DoScroll
    
    Call AddRoleButtonEnable
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
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmGuildRoleList.frm")
End Sub
Public Sub UpdateRoleList()
    lblGuildRolesQty.Caption = UBound(PlayerData.Guild.Roles) & "/" & PlayerData.Guild.MaxRoles
    Exit Sub
End Sub
Public Sub AddRoleButtonEnable()
    Dim AddRolePermission As Boolean
    
    AddRolePermission = HasPermission(GP_ROLE_CREATE_DELETE) And (UBound(PlayerData.Guild.Roles) < PlayerData.Guild.MaxRoles)
    ImgNewRole.Enabled = AddRolePermission
    Call cButtonNewRole.EnableButton(AddRolePermission)
    
    Exit Sub
End Sub
