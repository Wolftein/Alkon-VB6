VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Argentum Online"
   ClientHeight    =   7215
   ClientLeft      =   1950
   ClientTop       =   1515
   ClientWidth     =   5190
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7215
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Estado de los sockets"
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   4440
      Width           =   4815
      Begin VB.Label lblStatusProxy 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OFFLINE"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   3360
         TabIndex        =   18
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proxy:"
         Height          =   210
         Left            =   3405
         TabIndex        =   17
         Top             =   240
         Width           =   510
      End
      Begin VB.Label lblStatusMQ 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OFFLINE"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   795
         TabIndex        =   16
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Message Queue"
         Height          =   210
         Left            =   480
         TabIndex        =   15
         Top             =   240
         Width           =   1350
      End
   End
   Begin MSWinsockLib.Winsock sckProxySender 
      Left            =   840
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckMQReceiver 
      Index           =   0
      Left            =   840
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckMQListener 
      Left            =   360
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CheckBox chkServerHabilitado 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Server Habilitado Solo Gms"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5280
      Width           =   2775
   End
   Begin VB.TextBox txtNumUsers 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "0"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdSystray 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Systray"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdCerrarServer 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Cerrar Servidor"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6720
      Width           =   3495
   End
   Begin VB.CommandButton cmdConfiguracion 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Configuración General"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   4935
   End
   Begin VB.Timer tPiqueteC 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   3000
      Top             =   2580
   End
   Begin VB.Timer securityTimer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   3480
      Top             =   2100
   End
   Begin VB.CommandButton cmdDump 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Crear Log Crítico de Usuarios"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5760
      Width           =   4935
   End
   Begin VB.Timer FX 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   3960
      Top             =   2580
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   3060
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   3960
      Top             =   2100
   End
   Begin VB.Timer tLluviaEvent 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3480
      Top             =   3060
   End
   Begin VB.Timer tLluvia 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3480
      Top             =   2580
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3000
      Top             =   3060
   End
   Begin VB.Timer npcataca 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   4440
      Top             =   3060
   End
   Begin VB.Timer KillLog 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4440
      Top             =   2100
   End
   Begin VB.Timer TIMER_AI 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4455
      Top             =   2580
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Mensajea todos los clientes (Solo testeo)"
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4935
      Begin VB.Timer TimerDuelos 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3840
         Top             =   2880
      End
      Begin MSWinsockLib.Winsock sckStateServer 
         Left            =   240
         Top             =   1920
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer tMySQL 
         Interval        =   25000
         Left            =   2880
         Top             =   1380
      End
      Begin VB.TextBox txtChat 
         BackColor       =   &H00C0FFFF&
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   1320
         Width           =   4695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Enviar por Consola"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Enviar por Pop-Up"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox BroadMsg 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Label Escuch 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label CantUsuarios 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Número de usuarios jugando:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2460
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   15
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuCerrarBackup 
         Caption         =   "Cerrar (Con Backup)"
      End
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
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

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private iDay As Integer
    
Const NIM_ADD = 0
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONUP = &H205

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, ID As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
On Error GoTo ErrHandler
  
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = hWnd
    nidTemp.uID = ID
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function setNOTIFYICONDATA de frmMain.frm")
End Function

Private Sub Auditoria_Timer()
On Error GoTo errhand
    
    Call PasarSegundo 'sistema de desconexion de 10 segs
    
    
    Call SetSocketStatusLabels   ' Show the status of the sockets in frmMain
    
    ' Check tournaments
    If Tournament.CountdownActivated Then _
        Call TournamentCountdownCheck
    
    Call ActualizaEstadisticasWeb
    
    Exit Sub

errhand:

    Call LogError("Error en Timer Auditoria. Err: " & Err.Description & " - " & Err.Number)
    Resume Next

End Sub

Public Sub UpdateNpcsExp(ByVal Multiplicador As Single)
On Error GoTo ErrHandler
  

    Dim NpcIndex As Long
    For NpcIndex = 1 To LastNPC
        With Npclist(NpcIndex)
            .GiveEXP = .GiveEXP * Multiplicador
            .flags.ExpCount = .flags.ExpCount * Multiplicador
        End With
    Next NpcIndex
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateNpcsExp de frmMain.frm")
End Sub

Private Sub AutoSave_Timer()

On Error GoTo ErrHandler
'fired every minute
Static Minutos As Long
Static MinutosLatsClean As Long
Static MinsPjesSave As Long
Static MinsSendMotd As Long

Minutos = Minutos + 1
MinsPjesSave = MinsPjesSave + 1
MinsSendMotd = MinsSendMotd + 1

    Dim tmpHappyHour As Double
     
    ' HappyHour
    iDay = Weekday(Date)
    tmpHappyHour = HappyHourDays(iDay)
     
    If tmpHappyHour <> HappyHour Then
       
        If HappyHourActivated Then
            ' Reestablece la exp de los npcs
           If HappyHour <> 0 Then UpdateNpcsExp (1 / HappyHour)
         End If
       
        If tmpHappyHour = 1 Then ' Desactiva
           Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Ha concluido la Happy Hour!", FontTypeNames.FONTTYPE_DIOS))
             HappyHourActivated = False
       
        Else ' Activa
           UpdateNpcsExp tmpHappyHour
           
            If HappyHour <> 1 Then
                Call SendData(SendTarget.ToAll, 0, _
                    PrepareMessageConsoleMsg("Se ha modificado la Happy Hour, a partir de ahora las criaturas aumentan su experiencia en un " & Round((tmpHappyHour - 1) * 100, 2) & "%", FontTypeNames.FONTTYPE_DIOS))
            Else
                Call SendData(SendTarget.ToAll, 0, _
                    PrepareMessageConsoleMsg("¡Ha comenzado la Happy Hour! ¡Las criaturas aumentan su experiencia en un " & Round((tmpHappyHour - 1) * 100, 2) & "%!", FontTypeNames.FONTTYPE_DIOS))
            End If
           
             HappyHourActivated = True
        End If
     
        HappyHour = tmpHappyHour
    End If

If Minutos = MinutosWs - 1 Then
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Worldsave en 1 minuto ...", FontTypeNames.FONTTYPE_VENENO))
End If

If Minutos >= MinutosWs Then
    Call ES.DoBackUp
    Call aClon.VaciarColeccion
    Minutos = 0
End If

If MinsPjesSave = MinutosGuardarUsuarios - 1 Then
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("CharSave en 1 minuto ...", FontTypeNames.FONTTYPE_VENENO))
ElseIf MinsPjesSave >= MinutosGuardarUsuarios Then
    Call mdParty.ActualizaExperiencias
    Call GuardarUsuarios
    MinsPjesSave = 0
End If

If MinutosLatsClean >= 15 Then
    MinutosLatsClean = 0
    Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
    Call LimpiarMundo
Else
    MinutosLatsClean = MinutosLatsClean + 1
End If

'If MinsSendMotd >= MinutosMotd Then
'    Dim I As Long
'    For I = 1 To LastUser
'        If UserList(I).flags.UserLogged Then
'            Call SendMOTD(I)
'        End If
'    Next I
'    MinsSendMotd = 0
'End If

Call PurgarPenas

Call modSession.RemoveExpiredSessions

'<<<<<-------- Log the number of users online ------>>>
Dim N As Integer
N = FreeFile()
Open ServerConfiguration.LogsPaths.GeneralPath & "numusers.log" For Output Shared As N
Print #N, NumUsers
Close #N
'<<<<<-------- Log the number of users online ------>>>

Exit Sub
ErrHandler:
    Call LogError("Error en TimerAutoSave " & Err.Number & ": " & Err.Description)
    Resume Next
End Sub

Private Sub chkServerHabilitado_Click()
    ServerSoloGMs = chkServerHabilitado.value
  
End Sub

Private Sub cmdCerrarServer_Click()
    If MsgBox("¡¡Atencion!! Si cierra el servidor puede provocar la perdida de datos. " & _
        "¿Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
On Error GoTo ErrHandler
  
        Call General.StartShutDown(General.ShutDownFrom.User, False)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub cmdCerrarServer_Click de frmMain.frm")
End Sub

Private Sub cmdConfiguracion_Click()
    frmServidor.Visible = True
  
End Sub

Private Sub CMDDUMP_Click()
On Error Resume Next
On Error GoTo ErrHandler
  

    Dim I As Integer
    For I = 1 To MaxUsers
        Call LogCriticEvent(I & ") ConnidValida: " & UserList(I).ConnIDValida & " Name: " & UserList(I).Name & _
            " UserLogged: " & UserList(I).flags.UserLogged)
    Next I
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CMDDUMP_Click de frmMain.frm")
End Sub


Private Sub cmdSystray_Click()
    SetSystray
  
End Sub

Private Sub Command1_Click()
Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(BroadMsg.Text))
''''''''''''''''SOLO PARA EL TESTEO'''''''
''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
On Error GoTo ErrHandler
  
txtChat.Text = txtChat.Text & vbNewLine & "Servidor> " & BroadMsg.Text
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Command1_Click de frmMain.frm")
End Sub

Public Sub InitMain(ByVal f As Byte)
On Error GoTo ErrHandler
  

If f = 1 Then
    Call SetSystray
Else
    frmMain.Show
End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub InitMain de frmMain.frm")
End Sub

Private Sub Command2_Click()
Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & BroadMsg.Text, FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
''''''''''''''''SOLO PARA EL TESTEO'''''''
''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
On Error GoTo ErrHandler
  
txtChat.Text = txtChat.Text & vbNewLine & "Servidor> " & BroadMsg.Text
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Command2_Click de frmMain.frm")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
On Error GoTo ErrHandler
  
   
   If Not Visible Then
        Select Case X \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hWnd, hProcess
                AppActivate hProcess
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
        End Select
   End If
   
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_MouseMove de frmMain.frm")
End Sub

Public Sub QuitarIconoSystray()
On Error GoTo ErrHandler
  
On Error Resume Next

'Borramos el icono del systray
Dim I As Integer
Dim nid As NOTIFYICONDATA

nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

I = Shell_NotifyIconA(NIM_DELETE, nid)
    

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub QuitarIconoSystray de frmMain.frm")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandler
    If General.ShutdownBy = General.ShutDownFrom.Nobody Then
        Call General.StartShutDown(General.ShutDownFrom.System, True)
        Cancel = 1
    End If
    
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Unload de frmMain.frm")
End Sub

Private Sub FX_Timer()
On Error GoTo hayerror

Call SonidosMapas.ReproducirSonidosDeMapas

Exit Sub
hayerror:

End Sub

Private Sub GameTimer_Timer()
'********************************************************
'Author: Unknown
'Last Modify Date: -
'********************************************************
    Dim iUserIndex As Long
    Dim bEnviarStats As Boolean
    Dim bEnviarAyS As Boolean
    
On Error GoTo hayerror

    modIntervals.TickCount = GetTickCount()

    '<<<<<< Procesa eventos de los usuarios >>>>>>
    For iUserIndex = 1 To LastUser
        With UserList(iUserIndex)
           'Conexion activa?
           If .ConnIDValida Then
                '¿User valido?
                
                If .flags.UserLogged Then
                    
                    '[Alejo-18-5]
                    bEnviarStats = False
                    bEnviarAyS = False
                    
                    If .flags.Paralizado = 1 Then Call EfectoParalisisUser(iUserIndex)
                    
                    If .flags.Muerto = 0 Then
                        
                        If .flags.Putrefaccion Then Call EfectoPutrefaccion(iUserIndex)
                        
                        If (.flags.Privilegios And PlayerType.User) Then
                            If .flags.Desnudo Then Call EfectoFrio(iUserIndex)
                            
                            If .flags.Envenenado Then Call EfectoVeneno(iUserIndex)
                            If .flags.Ceguera Or .flags.Estupidez Then Call EfectoCegueEstu(iUserIndex)
                            If .flags.Petrificado Then Call EfectoPetrificado(iUserIndex)
                            
                            Call EfectoLava(iUserIndex)
                            Call DoHambre(iUserIndex, bEnviarAyS)
                            Call DoSed(iUserIndex, bEnviarAyS)
                        End If
                        
                        If .flags.Meditando Then Call DoMeditar(iUserIndex)
                        
                        If .flags.AdminInvisible <> 1 Then
                            If .flags.invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
                            If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex)
                            If .flags.Inmunidad > 0 Then Call EfectoInmunidad(iUserIndex)
                        End If
                        
                        If .flags.Mimetizado <> 0 Then Call EfectoMimetismo(iUserIndex)
                        
                        If .flags.AtacablePor <> 0 Then Call EfectoEstadoAtacable(iUserIndex)
                        
                        Call DuracionPociones(iUserIndex)
                        
                        If .flags.Hambre = 0 And .flags.Sed = 0 Then
                            Dim finishResting As Boolean

                            If Not .flags.Descansar Then
                            'No esta descansando
                                
                                If Not Lloviendo Or (Lloviendo And Not Intemperie(iUserIndex)) Then
                                    Call DoNormalHeal(iUserIndex, bEnviarStats)
                                    If bEnviarStats Then
                                        Call WriteUpdateHP(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    Call DoNormalStaminaRecovery(iUserIndex, bEnviarStats)
                                    If bEnviarStats Then
                                        Call WriteUpdateSta(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                End If
                            Else
                            'esta descansando
                                If HasPassiveAssigned(iUserIndex, ePassiveSpells.Regeneration) And Not finishResting Then
                                ' Then
                                    Call DoRegeneration(iUserIndex, bEnviarStats)
                                    finishResting = .Stats.MaxHp = .Stats.MinHp And .Stats.MaxSta = .Stats.MinSta And PassiveConditionMet(iUserIndex, ePassiveSpells.Regeneration)

                                Else
                                    Call DoRestingHeal(iUserIndex, bEnviarStats)
                                    finishResting = .Stats.MaxHp = .Stats.MinHp And .Stats.MaxSta = .Stats.MinSta
                                End If
                                
                                If bEnviarStats Then
                                    Call WriteUpdateHP(iUserIndex)
                                    bEnviarStats = False
                                End If
                                
                                Call DoRestingStaminaRecovery(iUserIndex, bEnviarStats)
                                If bEnviarStats Then
                                    Call WriteUpdateSta(iUserIndex)
                                    bEnviarStats = False
                                End If
                                'termina de descansar automaticamente
                                If finishResting Then
                                    Call WriteRestOK(iUserIndex)
                                    Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                    .flags.Descansar = False
                                End If
                                
                            End If
                        End If
                        
                        If bEnviarAyS Then Call WriteUpdateHungerAndThirst(iUserIndex)
                        
                        If .InvokedPetsCount > 0 Then Call TiempoInvocacion(iUserIndex)
                    Else
                        If .flags.Traveling <> 0 Then Call TravelingEffect(iUserIndex)
                    End If 'Muerto
                End If
                
                'Inactive players will be removed!
                If IsIntervalReached(.Counters.IdleCount) And Not .CraftingStore.IsOpen Then
                    If .flags.UserLogged Then
                        Call ExitSecureCommerce(iUserIndex)
                        Call Cerrar_Usuario(iUserIndex)
                    Else
                        Call CloseSocket(iUserIndex)
                    End If
                End If
            End If
        End With
    Next iUserIndex
Exit Sub

hayerror:
    LogError ("Error en GameTimer: " & Err.Description & " UserIndex = " & iUserIndex)
End Sub
  
Private Sub mnusalir_Click()
On Error GoTo ErrHandler
    Call cmdCerrarServer_Click
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub mnusalir_Click de frmMain.frm")
End Sub

Private Sub mnuCerrarBackup_Click()
    Call frmServidor.ApagarConBackup
  
End Sub

Public Sub mnuMostrar_Click()
On Error Resume Next
On Error GoTo ErrHandler
  
    WindowState = vbNormal
    Form_MouseMove 0, 0, 7725, 0
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub mnuMostrar_Click de frmMain.frm")
End Sub

Private Sub KillLog_Timer()
On Error Resume Next
On Error GoTo ErrHandler
  Dim LogPath As String
  LogPath = ServerConfiguration.LogsPaths.GeneralPath
If FileExist(LogPath & "connect.log", vbNormal) Then Kill LogPath & "connect.log"
If FileExist(LogPath & "haciendo.log", vbNormal) Then Kill LogPath & "haciendo.log"
If FileExist(LogPath & "stats.log", vbNormal) Then Kill LogPath & "stats.log"
If FileExist(LogPath & "Asesinatos.log", vbNormal) Then Kill LogPath & "Asesinatos.log"
If FileExist(LogPath & "HackAttemps.log", vbNormal) Then Kill LogPath & "HackAttemps.log"
If Not FileExist(LogPath & "nokillwsapi.txt") Then
    If FileExist(LogPath & "wsapi.log", vbNormal) Then Kill LogPath & "wsapi.log"
End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub KillLog_Timer de frmMain.frm")
End Sub

Private Sub SetSystray()
On Error GoTo ErrHandler
  

    Dim I As Integer
    Dim S As String
    Dim nid As NOTIFYICONDATA
    
    S = "ARGENTUM-ONLINE"
    nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, S)
    I = Shell_NotifyIconA(NIM_ADD, nid)
        
    If WindowState <> vbMinimized Then WindowState = vbMinimized
    Visible = False

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SetSystray de frmMain.frm")
End Sub

Private Sub sckProxySender_Close()
    sckProxySender.Close
    sckProxySender.Listen
End Sub

Private Sub sckProxySender_ConnectionRequest(ByVal requestID As Long)
On Error GoTo ErrHandler

    Call sckProxySender.Close
    Call sckProxySender.Accept(requestID)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub sckProxySender_ConnectionRequest de frmMain.frm")
End Sub

Private Sub sckProxySender_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call LogError("Error in sckProxySender. Number: " & Number & ", Description: " & Description)
    sckProxySender.Close
    sckProxySender.Listen
End Sub

Private Sub sckStateServer_Close()
    sckStateServer.Close
    sckStateServer.Listen
End Sub

Private Sub sckStateServer_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    sckStateServer.Close
    sckStateServer.Listen
End Sub

Private Sub securityTimer_Timer()

On Error GoTo ErrHandler
  
#If EnableSecurity Then
    Call Security.SecurityCheck
#End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub securityTimer_Timer de frmMain.frm")
End Sub

Private Sub TIMER_AI_Timer()

On Error GoTo ErrorHandler
    Dim NpcIndex As Long
    Dim mapa As Integer
    
    'Barrin 29/9/03
    If Not haciendoBK And Not EnPausa Then
        'Update NPCs
        For NpcIndex = 1 To LastNPC
                    
            With Npclist(NpcIndex)
                If .flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
                
                    ' Chequea si contiua teniendo dueño
                    If .Owner > 0 Then Call ValidarPermanenciaNpc(NpcIndex)
                
                    If .flags.Paralizado = 1 Then
                        Call EfectoParalisisNpc(NpcIndex)
                    Else
                        ' Preto? Tienen ai especial
                        If .NPCtype = eNPCType.Pretoriano Then
                            Call ClanPretoriano(.ClanIndex).PerformPretorianAI(NpcIndex)
                        Else
                            'Usamos AI si hay algun user en el mapa
                            If .flags.Inmovilizado = 1 Then
                               Call EfectoParalisisNpc(NpcIndex)
                            End If
                            
                            mapa = .Pos.Map
                            
                            If mapa > 0 Then
                                If MapInfo(mapa).NumUsers > 0 Then
                                    If .Movement <> TipoAI.ESTATICO Then
                                        If .Timers.Check(TimersIndex.Walk) Then
                                            Call .Timers.Restart(TimersIndex.MoveAttack)
                                            Call AIMovimiento(NpcIndex)
                                        End If
                                        Call AIAtaque(NpcIndex)
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    If .flags.Invocador > 0 Then Call CheckNpcInvocaciones(NpcIndex)
                End If
            End With
        Next NpcIndex
    End If
    
    Exit Sub

ErrorHandler:
    Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).Name & " mapa:" & _
    Npclist(NpcIndex).Pos.Map)
    Call MuereNpc(NpcIndex, 0)
End Sub

Private Sub TimerDuelos_Timer()
On Error GoTo ErrHandler
    Dim I As Byte

    For I = 1 To UBound(DuelData.Duelo)
        If Not DuelData.Duelo(I).estado = eDuelState.Vacio Then
            If DuelData.Duelo(I).estado = eDuelState.Esperando_Inicio Then
                If DuelData.Duelo(I).Counter > 0 Then
                    Call SendData(SendTarget.ToDuelo, I, PrepareMessageConsoleMsg("El duelo iniciará en: " & DuelData.Duelo(I).Counter & " segundos.", FontTypeNames.FONTTYPE_INFO))
                    DuelData.Duelo(I).Counter = DuelData.Duelo(I).Counter - 1
                    If DuelData.Duelo(I).Counter <= 0 Then
                        DuelData.Duelo(I).Counter = 0
                        Call IniciarDuelo(I)
                    End If
                End If
            End If
            If DuelData.Duelo(I).estado = eDuelState.Esperando_Final Then
                If DuelData.Duelo(I).Counter > 0 Then
                    DuelData.Duelo(I).Counter = DuelData.Duelo(I).Counter - 1
                    If DuelData.Duelo(I).Counter <= 0 Then
                        DuelData.Duelo(I).Counter = 0
                        Call CerrarDuelo(I)
                    End If
                End If
            End If
            If DuelData.Duelo(I).estado = eDuelState.Esperando_Jugadores Then
                If DuelData.Duelo(I).Counter > 0 Then
                    DuelData.Duelo(I).Counter = DuelData.Duelo(I).Counter - 1
                    If DuelData.Duelo(I).Counter <= 0 Then
                        DuelData.Duelo(I).Counter = 0
                        Call CancelarDuelo(I, False)
                    End If
                End If
            End If
            If DuelData.Duelo(I).estado = eDuelState.Iniciado Then
                DuelData.Duelo(I).Counter = DuelData.Duelo(I).Counter - 1
                If DuelData.Duelo(I).Counter = 0 Then
                    
                    Call TerminarDueloTimeout(I)
                End If
                
                
            End If
        End If
    Next I
    
    Exit Sub
ErrHandler:
        Call LogError("Error en TimerRetos." & " " & Err.Description)
End Sub

Private Sub tLluvia_Timer()
On Error GoTo ErrHandler

Dim iCount As Long
If Lloviendo Then
   For iCount = 1 To LastUser
        Call EfectoLluvia(iCount)
   Next iCount
End If

Exit Sub
ErrHandler:
Call LogError("tLluvia " & Err.Number & ": " & Err.Description)
End Sub

Private Sub tLluviaEvent_Timer()

On Error GoTo ErrorHandler
Static MinutosLloviendo As Long
Static MinutosSinLluvia As Long

If Not Lloviendo Then
    MinutosSinLluvia = MinutosSinLluvia + 1
    If MinutosSinLluvia >= 15 And MinutosSinLluvia < 1440 Then
        If RandomNumber(1, 100) <= 2 Then
            Lloviendo = True
            MinutosSinLluvia = 0
            Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
        End If
    ElseIf MinutosSinLluvia >= 1440 Then
        Lloviendo = True
        MinutosSinLluvia = 0
        Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
    End If
Else
    MinutosLloviendo = MinutosLloviendo + 1
    If MinutosLloviendo >= 5 Then
        Lloviendo = False
        Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
        MinutosLloviendo = 0
    Else
        If RandomNumber(1, 100) <= 2 Then
            Lloviendo = False
            MinutosLloviendo = 0
            Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
        End If
    End If
End If

Exit Sub
ErrorHandler:
Call LogError("Error tLluviaTimer")

End Sub

Private Sub tMySQL_Timer()
    Call ExecuteSql("SELECT 1")
  
End Sub

Private Sub tPiqueteC_Timer()
   
    Dim I As Long
    
On Error GoTo ErrHandler
    For I = 1 To LastUser
        With UserList(I)
            If .flags.UserLogged And Not EsGm(I) Then
                If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.ANTIPIQUETE Then
                    If .flags.Muerto = 0 Then
                        .Counters.PiqueteC = .Counters.PiqueteC + 1
                        Call WriteConsoleMsg(I, "¡¡¡Estás obstruyendo la vía pública, muévete o serás encarcelado!!!", FontTypeNames.FONTTYPE_INFO)
                        
                        If .Counters.PiqueteC > 23 Then
                            .Counters.PiqueteC = 0
                            Call Encarcelar(I, Constantes.TiempoCarcelPiquete)
                        End If
                    Else
                        .Counters.PiqueteC = 0
                    End If
                Else
                    .Counters.PiqueteC = 0
                End If
            End If
        End With
    Next I
Exit Sub

ErrHandler:
    Call LogError("Error en tPiqueteC_Timer " & Err.Number & ": " & Err.Description)
End Sub

Private Sub sckMQListener_ConnectionRequest(ByVal requestID As Long)
    Dim I As Integer
On Error GoTo ErrHandler
  
    I = getFreeSocket()
        
    If I <> -1 Then
        Call sckMQReceiver.Item(I).Accept(requestID)
        Debug.Print "Connection accepted: " & requestID & "(" & CStr(I) & ")"
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub sckMQListener_ConnectionRequest de frmMain.frm")
End Sub

Private Sub sckMQReceiver_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    On Error GoTo ErrHandler:
    
    Call ProcessMQSocketData(bytesTotal)
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error en sckMQReceiver_DataArrival. Err: " & Err.Number & " - " & Err.Description)
End Sub

Private Function ProcessMQSocketData(ByVal bytesTotal) As Boolean
    
    Dim strData As String
    Dim receivedData() As Byte
    
    ReDim receivedData(0 To bytesTotal - 1)
    ' Get the received data as a byte array
    sckStateServer.GetData receivedData
    
    ' Insert the received data into the MQ Received buffer
    Call modMessageQueueProxy.MQReceivedDataBuffer.WriteBlock(receivedData)

    ' If the socket sent more than one message at the same time
    ' then read the list of received messages.
    'Dim msgSpltd() As String
    'msgSpltd = Split(strData, "<EOF>")
    
    Dim keepProcessing As Boolean: keepProcessing = False
        
    Do
    
        keepProcessing = modMessageQueueProxy.HandleMQReceiverMessage(modMessageQueueProxy.MQReceivedDataBuffer)
    
    Loop Until keepProcessing = False

End Function

Private Sub sckStateServer_ConnectionRequest(ByVal requestID As Long)
On Error GoTo ErrHandler

    Dim I As Integer
    
    Call sckStateServer.Close
    Call sckStateServer.Accept(requestID)
    
    Call modStateServer.OnConnected
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub sckMQListener_ConnectionRequest de frmMain.frm")
End Sub


Private Sub sckStateServer_DataArrival(ByVal bytesTotal As Long)
On Error GoTo ErrHandler

    Dim strData As String
    Dim receivedData() As Byte
    
    ReDim receivedData(0 To bytesTotal - 1)
    ' Get the received data as a byte array
    sckStateServer.GetData receivedData
    
    ' Insert the received data into the MQ Received buffer
    Call modStateServer.InboundByteQueue.WriteBlock(receivedData)

    Dim keepProcessing As Boolean: keepProcessing = False
        
    Do
        keepProcessing = modStateServer.HandleStateServerMessage(modStateServer.InboundByteQueue)
    Loop Until keepProcessing = False

  
     
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub sckStateServer_DataArrival de frmMain.frm")
End Sub

Private Sub sckMQReceiver_Close(Index As Integer)
    sckMQReceiver(Index).Close
On Error GoTo ErrHandler
  
    Debug.Print "Closed " & CStr(Index)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub sckMQReceiver_Close de frmMain.frm")
End Sub

Private Function getFreeSocket() As Integer
On Error GoTo ErrHandler
  
    Dim I As Integer
    I = 0
    
    getFreeSocket = -1
    
    For I = 0 To sckMQReceiver.Count - 1
        If sckMQReceiver(I).State = sckClosed Then
            getFreeSocket = I
            Exit For
        End If
    Next I

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function getFreeSocket de frmMain.frm")
End Function

Public Sub ListenMQ()
On Error GoTo ErrHandler
  
    sckMQListener.LocalPort = 9999
    sckMQListener.Listen
    
    Load sckMQReceiver(1)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ListenMQ de frmMain.frm")
End Sub

Public Sub ListenProxySender()
On Error GoTo Err:

    'sckProxySender.RemoteHost = "127.0.0.1"
    'sckProxySender.RemotePort = 45054
    sckProxySender.LocalPort = 45054
    sckProxySender.Protocol = sckTCPProtocol
    
    sckProxySender.Close
    
    Call sckProxySender.Listen
    
    Exit Sub
    
Err:
    Debug.Print Err.Description
End Sub

Public Sub SetSocketStatusLabels()
On Error GoTo ErrHandler
  
    Dim I As Integer
    Dim mqConnected As Boolean
    mqConnected = False
    For I = 0 To sckMQReceiver.Count - 1
        If sckMQReceiver.Item(I).State = sckConnected Then
            mqConnected = True
        End If
    Next I
    

    If mqConnected Then
        frmMain.lblStatusMQ.Caption = "ONLINE"
        frmMain.lblStatusMQ.ForeColor = &HC000&
    Else
        frmMain.lblStatusMQ.Caption = "OFFLINE"
        frmMain.lblStatusMQ.ForeColor = &HFF&
    End If


    If frmMain.sckProxySender.State = sckConnected Then
        frmMain.lblStatusProxy.Caption = "ONLINE"
        frmMain.lblStatusProxy.ForeColor = &HC000&
    Else
        frmMain.lblStatusProxy.Caption = "OFFLINE"
        frmMain.lblStatusProxy.ForeColor = &HFF&
    End If


  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SetSocketStatusLabels de frmMain.frm")
End Sub


Public Sub ListenForRemoteTools()
    
    With ServerConfiguration.ExternalTools
        
        If ServerConfiguration.ExternalTools.StateServer.Enabled Then
            frmMain.sckStateServer.Close
            
            sckStateServer.LocalPort = ServerConfiguration.ExternalTools.StateServer.ListenPort
            sckStateServer.Protocol = sckTCPProtocol
            
            Call sckStateServer.Listen
        End If
        
        If ServerConfiguration.ExternalTools.ProxyServer.Enabled Then
            frmMain.sckProxySender.Close
            
            sckProxySender.LocalPort = ServerConfiguration.ExternalTools.ProxyServer.ListenPort
            sckProxySender.Protocol = sckTCPProtocol
            
            Call sckProxySender.Listen
        End If

    End With
    
End Sub
