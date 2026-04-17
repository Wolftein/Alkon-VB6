VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGuildMsgBoxHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsMessageBoxHandler

Private Sub clsMessageBoxHandler_OnAccept()

End Sub

Private Sub clsMessageBoxHandler_OnNo()
    Call Guilds.GuildRejectInvitation
End Sub

Private Sub clsMessageBoxHandler_OnYes()
    Call Guilds.GuildAcceptInvitation
End Sub
