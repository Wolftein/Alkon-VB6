VERSION 5.00
Begin VB.Form frmGuildManagement 
   Caption         =   "Administracion de clanes"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8715
   Icon            =   "frmGuildManagement.frx":0000
   ScaleHeight     =   5595
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExpellMembers 
      Caption         =   "Expulsar Miembros"
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmdReloadFromDB 
      Caption         =   "Recargar desde DB"
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Frame frmActions 
      Caption         =   "Acciones"
      Height          =   3135
      Left            =   5640
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton btnSaveAlignment 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   435
         Width           =   735
      End
      Begin VB.ComboBox cboAlign 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblError 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Alineacion"
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   5040
      Width           =   2055
   End
   Begin VB.ListBox lstMembers 
      Height          =   3960
      Left            =   3000
      TabIndex        =   3
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton cmdLoadGuilds 
      Caption         =   "Leer Guilds"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   7935
   End
   Begin VB.ListBox lstGuilds 
      Height          =   3960
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label lblMiembros 
      Caption         =   "Miembros"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Clanes"
      Height          =   195
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   480
   End
End
Attribute VB_Name = "frmGuildManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnSaveAlignment_Click()
    Dim selectedGuild As Integer
On Error GoTo ErrHandler
  
    selectedGuild = lstGuilds.ItemData(lstGuilds.ListIndex)
    
    If selectedGuild > 0 And cboAlign.ListIndex > 0 Then
        ' Assign the new Alignment to the guild
        'modGuilds.Guilds(selectedGuild).Alignment = cboAlign.ListIndex
    Else
        lblError.Caption = "Seleccion inválida"
    End If
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub btnSaveAlignment_Click de frmGuildManagement.frm")
End Sub

Private Sub cmdClose_Click()
    Unload Me

End Sub

Private Sub cmdLoadGuilds_Click()
    On Error GoTo Err:
        Dim I As Integer
        lstGuilds.Clear
        lstMembers.Clear
        
        'For I = 1 To modGuilds.NroClanes
        '    Call lstGuilds.AddItem(modGuilds.Guilds(I).Name)
        '    lstGuilds.ItemData(lstGuilds.NewIndex) = I
        'Next I
    
    Exit Sub
Err:
    Debug.Print (Err.Description)
    ' Error
End Sub

Private Sub Form_Load()
    Call cboAlign.AddItem("-----", 0)
On Error GoTo ErrHandler
  
    Call cboAlign.AddItem("Legión", 1)
    Call cboAlign.AddItem("Criminal", 2)
    Call cboAlign.AddItem("Neutral", 3)
    Call cboAlign.AddItem("Ciudadano", 4)
    Call cboAlign.AddItem("Armada", 5)
    Call cboAlign.AddItem("GameMaster", 6)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmGuildManagement.frm")
End Sub

Private Sub lstGuilds_Click()
On Error GoTo Err:

    Dim I As Integer
    Dim selectedGuildIndex As Integer
    Debug.Print lstGuilds.Text
    Debug.Print lstGuilds.ItemData(lstGuilds.ListIndex)
    Dim guildMembers() As String
    selectedGuildIndex = CInt(lstGuilds.ItemData(lstGuilds.ListIndex))
    lstMembers.Clear
    
    'If modGuilds.Guilds(selectedGuildIndex).GetTotalMembers() > 0 Then
    '    guildMembers = modGuilds.Guilds(selectedGuildIndex).GetMemberList()

    '    For I = 1 To UBound(guildMembers)
    '       Call lstMembers.AddItem(guildMembers(I))
    '    Next I
        
    '    cboAlign.ListIndex = modGuilds.Guilds(selectedGuildIndex).Alignment
    '    frmActions.Visible = True
    'Else
    '    frmActions.Visible = False
    'End If
        
    
    
    Exit Sub
Err:
    Debug.Print (Err.Description)
End Sub
