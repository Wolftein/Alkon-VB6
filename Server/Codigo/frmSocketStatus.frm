VERSION 5.00
Begin VB.Form frmSocketStatus 
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2565
   Icon            =   "frmSocketStatus.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   2565
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Estado de los sockets"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.Timer tmrSocketStatus 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin VB.Label Label5 
         Caption         =   "Proxy Server:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblProxy 
         BackColor       =   &H0000C000&
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblMQ 
         BackColor       =   &H000000FF&
         Caption         =   "         "
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblStateServer 
         BackColor       =   &H0000C000&
         Caption         =   "         "
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "MQ Server:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "State Server:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   930
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C000&
      Caption         =   "         "
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "frmSocketStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    tmrSocketStatus.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrSocketStatus.Enabled = False

End Sub

Private Sub tmrSocketStatus_Timer()
     If frmMain.sckStateServer.State = sckConnected Then
On Error GoTo ErrHandler
  
        lblStateServer.BackColor = &HC000&
    Else
        lblStateServer.BackColor = &HFF&
    End If
    
    If frmMain.sckProxySender.State = sckConnected Then
        lblProxy.BackColor = &HC000&
    Else
        lblProxy.BackColor = &HFF&
    End If
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub tmrSocketStatus_Timer de frmSocketStatus.frm")
End Sub
