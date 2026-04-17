VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form frmSessionsManagement 
   Caption         =   "Manejo de Sesiones"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRemoveSessions 
      Caption         =   "X"
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   4200
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Session Data"
      Height          =   5055
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   5895
      Begin VB.ListBox lstCharacters 
         Height          =   1425
         Left            =   360
         TabIndex        =   11
         Top             =   2880
         Width           =   5055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Characters in session"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   2520
         Width           =   1500
      End
      Begin VB.Label lblAccount 
         Caption         =   "asdasd"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lblSessionToken 
         Caption         =   "TOKEN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Token:"
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   1320
         Width           =   510
      End
      Begin VB.Label lblSessionId 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   1200
         TabIndex        =   7
         Top             =   600
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Session ID:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Account:"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Session List"
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin MSComctlLib.ListView lstSessions 
         Height          =   3255
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   5741
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.CommandButton cmdReload 
         Caption         =   "Recargar Sesiones"
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label lblSesiones 
         AutoSize        =   -1  'True
         Caption         =   "Session amount: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmSessionsManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'La idea de este formulario es manejar las sesiones activas y agregar mas
'para simular grandes cantidades de trafico ficticio.

Option Explicit

Private NumSessions As Integer

Private Sub cmdReload_Click()

    Call ReloadSessions
    
End Sub

Private Sub cmdRemoveSessions_Click()
    Dim I As Integer
    For I = 0 To modSession.GetMaxAllowedSessions()
        modSession.CleanSessionSlot (I)
    Next I
End Sub

Private Sub Form_Load()

On Error GoTo ErrHandler
  
    Call ReloadSessions
  
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmSessionsManagement.frm")
End Sub

Private Sub ReloadSessions()
On Error GoTo ErrHandler
  
    Dim I As Integer
    
    Dim Item As ListItem
    lstSessions.View = lvwReport
    
    NumSessions = 0
    lstSessions.ListItems.Clear
    lstSessions.Appearance = ccFlat
    lstSessions.ColumnHeaders.Clear
    lstSessions.ColumnHeaders.Add , , "ID"
    lstSessions.ColumnHeaders.Add , , "C.Code"
    lstSessions.ColumnHeaders.Add , , "S.TOKEN"

    For I = 0 To modSession.GetMaxAllowedSessions()
        With aActiveSessions(I)
            'If .ServerTempCode <> vbNullString Then
                Set Item = lstSessions.ListItems.Add(NumSessions + 1, "ID" & I, I)
                Item.SubItems(1) = .ClientTempCode
                Item.SubItems(2) = .ServerTempCode
                
                NumSessions = NumSessions + 1
            'End If
        End With
        
    Next I
    
    Call lstSessions.Refresh
    
    lblSesiones.Caption = "Session amount: " & NumSessions
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ReloadSessions de frmSessionsManagement.frm")
End Sub

Private Sub lstSessions_Click()
    Dim SessionId As Integer
    Dim I As Integer
    
    If lstSessions.SelectedItem Is Nothing Then Exit Sub
    
    SessionId = CInt(lstSessions.SelectedItem.Text)
        
    With aActiveSessions(SessionId)
        lblSessionId.Caption = SessionId
        lblSessionToken.Caption = .Token
        lblAccount.Caption = .sAccountName
        
        lstCharacters.Clear
        
        For I = 1 To 8
            If .asAccountCharNames(I).CharName <> vbNullString Then
                Call lstCharacters.AddItem(.asAccountCharNames(I).CharId & " - " & .asAccountCharNames(I).CharName)
            End If
        Next I
    End With
    
End Sub

