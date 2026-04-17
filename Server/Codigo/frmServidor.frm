VERSION 5.00
Begin VB.Form frmServidor 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Configuración del Servidor"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   490
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReloadGuildReservedNames 
      BackColor       =   &H00FFC0C0&
      Caption         =   "GuildNam. Res."
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuests 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Quests"
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdReloadProfessions 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Professions.dat"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdSessionsMenu 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sesiones"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton cmdSocketStatus 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Stat de Sockets"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton cmdReiniciar 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Reiniciar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Administración"
      Height          =   2895
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   6375
      Begin VB.CommandButton cmdGuildMenu 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Lista de Clanes"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton cmdResetListen 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Reset Listen"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton cmdResetSockets 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Reset sockets"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton cmdDebugUserlist 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Debug UserList"
         Height          =   375
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton cmdUnbanAllIps 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Unban All IPs (PELIGRO!)"
         Height          =   495
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton cmdDebugNpcs 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Debug Npcs"
         Height          =   495
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton frmAdministracion 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Administración"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdPausarServidor 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Pausar el servidor"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdStatsSlots 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Stats de Slots"
         Height          =   375
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdVerTrafico 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Tráfico"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdConfigIntervalos 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Config. Intervalos"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdCerrar 
      BackColor       =   &H00FFC0C0&
      Cancel          =   -1  'True
      Caption         =   "Salir (Esc)"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdApagarServer 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Apagar Server con Backup"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Backup"
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   5880
      Width           =   6375
      Begin VB.CommandButton cmdLoadWorldBackup 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cargar Mapas"
         Height          =   375
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdCharBackup 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Guardar Chars"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdWorldBackup 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Guardar Mapas"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Recargar"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton cmdReloadGuilds 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Guilds"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdRecargarAdministradores 
         BackColor       =   &H0080C0FF&
         Caption         =   "Adminis"
         Height          =   375
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdRecargarGuardiasPosOrig 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Guardias en pos originales"
         Height          =   375
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1320
         Width           =   3015
      End
      Begin VB.CommandButton cmdRecargarMOTD 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MOTD"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton cmdRecargarMD5s 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MD5s"
         Height          =   375
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdRecargarServerIni 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Server.ini"
         Height          =   375
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdRecargarNombresInvalidos 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nomb Inval."
         Height          =   375
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdRecargarNPCs 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Npcs.dat"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdRecargarBalance 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Balance.dat"
         Height          =   375
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdRecargarHechizos 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Hechizos.dat"
         Height          =   375
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdRecargarObjetos 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Obj.dat"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmServidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.14.0
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Private Sub cmdGuildMenu_Click()
    frmGuildManagement.Show
End Sub

Private Sub cmdQuests_Click()
    Call LoadGuildQuests
    Call ReloadCurrentQuests
End Sub

Private Sub cmdReloadGuildReservedNames_Click()
    Call modGuild_Functions.LoadReservedNames
End Sub

Private Sub cmdReloadGuilds_Click()
    Call LoadGuilds
End Sub

Private Sub cmdReloadProfessions_Click()
On Error GoTo ErrHandler
  
    Call LoadProfessions
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub cmdReloadProfessions_Click de frmServidor.frm")
End Sub

Private Sub cmdSessionsMenu_Click()
    Call frmSessionsManagement.Show
End Sub

Private Sub cmdSocketStatus_Click()
    frmSocketStatus.Show
End Sub


Private Sub Form_Load()
    cmdResetSockets.Visible = True
On Error GoTo ErrHandler
  
    cmdResetListen.Visible = True
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmServidor.frm")
End Sub

Private Sub cmdApagarServer_Click()
    ApagarConBackup
  
End Sub

Public Sub ApagarConBackup()
On Error GoTo ErrHandler
  
    Dim N As Integer
    
    If MsgBox("¿Está seguro que desea hacer WorldSave, guardar pjs y cerrar?", vbYesNo, _
        "Apagar Magicamente") = vbNo Then Exit Sub
    
    Call General.StartShutDown(General.ShutDownFrom.User, True)
    Me.MousePointer = 11
    
    FrmStat.Show
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ApagarConBackup de frmServidor.frm")
End Sub

Private Sub cmdCerrar_Click()
    frmServidor.Visible = False
  
End Sub

Private Sub cmdCharBackup_Click()
    Me.MousePointer = 11
On Error GoTo ErrHandler
  
    Call mdParty.ActualizaExperiencias
    Call GuardarUsuarios
    Me.MousePointer = 0
    MsgBox "Grabado de personajes OK!"
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub cmdCharBackup_Click de frmServidor.frm")
End Sub

Private Sub cmdConfigIntervalos_Click()
    FrmInterv.Show
  
End Sub

Private Sub cmdDebugNpcs_Click()
    frmDebugNpc.Show

End Sub

Private Sub cmdDebugUserlist_Click()
    frmUserList.Show

End Sub

Private Sub cmdLoadWorldBackup_Click()
'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error GoTo ErrHandler
  
On Error Resume Next

    If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."
    
    Dim LogPath As String
    
    LogPath = ServerConfiguration.LogsPaths.GeneralPath
    FrmStat.Show
    
    If FileExist(LogPath & "errores.log", vbNormal) Then Kill LogPath & "errores.log"
    If FileExist(LogPath & "connect.log", vbNormal) Then Kill LogPath & "Connect.log"
    If FileExist(LogPath & "HackAttemps.log", vbNormal) Then Kill LogPath & "HackAttemps.log"
    If FileExist(LogPath & "Asesinatos.log", vbNormal) Then Kill LogPath & "Asesinatos.log"
    If FileExist(LogPath & "Resurrecciones.log", vbNormal) Then Kill LogPath & "Resurrecciones.log"
    If FileExist(LogPath & "Teleports.Log", vbNormal) Then Kill LogPath & "Teleports.Log"

    Call TCP.Disconnect
    
    Dim LoopC As Integer
    
    For LoopC = 1 To MaxUsers
        Call CloseSocket(LoopC)
    Next LoopC
    
    LastUser = 0
    NumUsers = 0
    
    Call FreeNPCs
    Call FreeCharIndexes
    
    Call LoadSini
    Call CargarBackUp
    Call LoadOBJData

    Call TCP.Listen("0.0.0.0", CStr(Puerto))
    
    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub cmdLoadWorldBackup_Click de frmServidor.frm")
End Sub

Private Sub cmdPausarServidor_Click()
    If EnPausa = False Then
On Error GoTo ErrHandler
  
        EnPausa = True
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
        cmdPausarServidor.Caption = "Reanudar el servidor"
    Else
        EnPausa = False
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
        cmdPausarServidor.Caption = "Pausar el servidor"
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub cmdPausarServidor_Click de frmServidor.frm")
End Sub

Private Sub cmdRecargarBalance_Click()
    Call LoadBalance
  
End Sub

Private Sub cmdRecargarGuardiasPosOrig_Click()
    Call ReSpawnOrigPosNpcs

End Sub

Private Sub cmdRecargarHechizos_Click()
    Call CargarHechizos

End Sub

Private Sub cmdRecargarMD5s_Click()
    Call MD5sCarga

End Sub

Private Sub cmdRecargarMOTD_Click()
    Call LoadMotd

End Sub

Private Sub cmdRecargarNombresInvalidos_Click()
    Call CargarForbidenWords

End Sub

Private Sub cmdRecargarNPCs_Click()
    Call CargaNpcsDat

End Sub

Private Sub cmdRecargarObjetos_Click()
    Call ResetForums
On Error GoTo ErrHandler
  
    Call LoadOBJData
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub cmdRecargarObjetos_Click de frmServidor.frm")
End Sub

Private Sub cmdRecargarServerIni_Click()
    Call LoadSini

End Sub

Private Sub cmdReiniciar_Click()

On Error GoTo ErrHandler
  
    If MsgBox("¡¡Atencion!! Si reinicia el servidor puede provocar la pérdida de datos de los usarios. " & _
    "¿Desea reiniciar el servidor de todas maneras?", vbYesNo) = vbNo Then Exit Sub
    
    Me.Visible = False
    Call General.Restart

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub cmdReiniciar_Click de frmServidor.frm")
End Sub

Private Sub cmdResetListen_Click()
    'Cierra el socket de escucha
On Error GoTo ErrHandler
  
    Call TCP.Disconnect
    Call TCP.Listen("0.0.0.0", CStr(Puerto))
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub cmdResetListen_Click de frmServidor.frm")
End Sub

Private Sub cmdResetSockets_Click()
    If MsgBox("¿Está seguro que desea reiniciar los sockets? Se cerrarán todas las conexiones activas.", vbYesNo, "Reiniciar Sockets") = vbYes Then
On Error GoTo ErrHandler
  
        'WOLFTEIN
    'Call WSApiReiniciarSockets
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub cmdResetSockets_Click de frmServidor.frm")
End Sub

Private Sub cmdStatsSlots_Click()
    frmConID.Show

End Sub

Private Sub cmdUnbanAllIps_Click()
    Dim I As Long, N As Long
On Error GoTo ErrHandler
  
    
    Dim sENtrada As String
    
    sENtrada = InputBox("Escribe ""estoy DE acuerdo"" sin comillas y con distinción de mayúsculas minúsculas para desbanear a todos los personajes", "UnBan", "hola")
    If sENtrada = "estoy DE acuerdo" Then
        
        N = BanIps.Count
        For I = 1 To BanIps.Count
            BanIps.Remove 1
        Next I
        
        MsgBox "Se han habilitado " & N & " ipes"
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub cmdUnbanAllIps_Click de frmServidor.frm")
End Sub

Private Sub cmdVerTrafico_Click()
    frmTrafic.Show

End Sub

Private Sub cmdWorldBackup_Click()
On Error GoTo ErrHandler

    Me.MousePointer = 11
    FrmStat.Show
    Call ES.DoBackUp
    Me.MousePointer = 0
    MsgBox "WORLDSAVE OK!!"
    
    Exit Sub

ErrHandler:
    Call LogError("Error en WORLDSAVE")
End Sub

Private Sub Form_Deactivate()
    frmServidor.Visible = False

End Sub

Private Sub frmAdministracion_Click()
    Me.Visible = False
On Error GoTo ErrHandler
  
    frmAdmin.Show
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub frmAdministracion_Click de frmServidor.frm")
End Sub

Private Sub cmdRecargarAdministradores_Click()
    loadAdministrativeUsers

End Sub

