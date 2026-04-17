VERSION 5.00
Begin VB.Form frmGuildEditRoles 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pboRolesPermissions 
      BorderStyle     =   0  'None
      Height          =   3700
      Left            =   0
      ScaleHeight     =   3705
      ScaleWidth      =   7200
      TabIndex        =   1
      Top             =   2300
      Width           =   7200
   End
   Begin VB.PictureBox pboRolesList 
      BorderStyle     =   0  'None
      Height          =   2300
      Left            =   0
      ScaleHeight     =   3000
      ScaleMode       =   0  'User
      ScaleWidth      =   7335
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmGuildEditRoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CloseWindow
End Sub

Private Sub Form_Load()

    Call frmGuildRolesList.Show
    
    SetParent frmGuildRolesList.hwnd, pboRolesList.hwnd
    frmGuildRolesList.Move 0, 0, pboRolesList.Width, pboRolesList.Height
    frmGuildRolesList.Visible = True
    
    Call modCustomCursors.SetFormCursorDefault(Me)
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Unload frmGuildRolesList
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
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmGuildEditRoles.frm")
End Sub
