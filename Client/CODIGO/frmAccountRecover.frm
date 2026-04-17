VERSION 5.00
Begin VB.Form frmAccountRecover 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Recuperar cuenta"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2355
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   2355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame f_Recover 
      Caption         =   "Recuperar"
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
      Begin VB.TextBox txtToken 
         Height          =   375
         Left            =   120
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton cmdRecover 
         Caption         =   "Recuperar"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   120
         MaxLength       =   30
         TabIndex        =   1
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Token:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmAccountRecover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub cmdRecover_Click()

On Error GoTo ErrHandler
  
If Not MainTimer.Check(TimersIndex.Action) Then Exit Sub

If Trim(frmAccountRecover.txtName.text) = "" Or Trim(frmAccountRecover.txtToken.text) = "" Then
    Call MsgBox("Los campos Cuenta y Token no pueden estar vacíos")
    Exit Sub
End If

    Call modAccount.Set_Acc_Data_To_Recover

    Call modAccount.Prepare_And_Connect(E_MODO.AccountRecover)

  
  Exit Sub
 
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub cmdRecover_Click de frmAccountRecover.frm")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CloseWindow
End Sub

Private Sub CloseWindow()
On Error GoTo ErrHandler
  
    Unload Me
    frmConnect.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmAccountRecover.frm")
End Sub

Private Sub Form_Load()
    Call modCustomCursors.SetFormCursorDefault(Me)
End Sub
