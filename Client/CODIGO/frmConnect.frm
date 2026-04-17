VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.TextBox txtPasswd 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   3960
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4620
      Width           =   4095
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3960
      TabIndex        =   0
      Top             =   3660
      Width           =   4095
   End
   Begin VB.TextBox PortTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4350
      TabIndex        =   2
      Text            =   "7666"
      Top             =   2700
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   5550
      TabIndex        =   4
      Text            =   "localhost"
      Top             =   2700
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.Image imgWebsiteLink 
      Height          =   615
      Left            =   3960
      Top             =   8160
      Width           =   4455
   End
   Begin VB.Image imgConectarse 
      Appearance      =   0  'Flat
      Height          =   540
      Left            =   6555
      Top             =   5640
      Width           =   1245
   End
   Begin VB.Image imgDiscord 
      Height          =   555
      Left            =   11160
      Top             =   8205
      Width           =   630
   End
   Begin VB.Image imgSalir 
      Appearance      =   0  'Flat
      Height          =   540
      Left            =   5370
      Top             =   6345
      Width           =   1245
   End
   Begin VB.Image imgCodigoFuente 
      Height          =   525
      Left            =   360
      Top             =   6120
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.Image imgReglamento 
      Height          =   525
      Left            =   360
      Top             =   4320
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.Image imgManual 
      Height          =   525
      Left            =   360
      Top             =   5280
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.Image imgRecuperar 
      Height          =   780
      Left            =   2880
      Top             =   6960
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Image imgCrearCuenta 
      Appearance      =   0  'Flat
      Height          =   540
      Left            =   4200
      Top             =   5640
      Width           =   1245
   End
   Begin VB.Image imgServArgentina 
      Height          =   795
      Left            =   360
      MousePointer    =   99  'Custom
      Top             =   9240
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   630
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Matías Fernando Pequeño
'matux@fibertel.com.ar
'www.noland-studios.com.ar
'Acoyte 678 Piso 17 Dto B
'Capital Federal, Buenos Aires - Republica Argentina
'Código Postal 1405

Option Explicit

Private cBotonCrearCuenta As clsGraphicalButton
Private cBotonRecuperarPass As clsGraphicalButton
Private cButtonWebsite As clsGraphicalButton
Private cBotonReglamento As clsGraphicalButton
Private cBotonCodigoFuente As clsGraphicalButton
Private cBotonBorrarPj As clsGraphicalButton
Private cBotonSalir As clsGraphicalButton
Private cBotonLeerMas As clsGraphicalButton
Private cButtonDiscord As clsGraphicalButton
Private cBotonConectarse As clsGraphicalButton


Public GettingPort As Boolean
Public sPuertoActual As String
Public LastButtonPressed As clsGraphicalButton

Private clsFormulario As clsFormMovementManager

Private Sub Form_Activate()
    'On Error Resume Next
On Error GoTo ErrHandler
  
    If ServersRecibidos Then
        If CurServer <> 0 Then
            IPTxt = ServersLst(1).Ip
            PortTxt = ServersLst(1).Puerto
        Else
            IPTxt = IPdelServidor
            PortTxt = PuertoDelServidor
        End If
    End If
        
    If (GameConfig.Extras.Name <> vbNullString) Then
        txtNombre = GameConfig.Extras.Name
        
        If txtPasswd.Visible Then
            Call txtPasswd.SetFocus
        End If
        
    Else
        Call txtNombre.SetFocus
    End If
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Activate de frmConnect.frm")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then prgRun = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'Make Server IP and Port box visible
On Error GoTo ErrHandler
  
If KeyCode = vbKeyI And Shift = vbCtrlMask Then
    
    'Port
    PortTxt.Visible = True
    'Label4.Visible = True
    
    'Server IP
    PortTxt.text = "7666"
    IPTxt.text = "127.0.0.1"
    IPTxt.Visible = True
    'Label5.Visible = True
    
    KeyCode = 0
    Exit Sub
End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_KeyUp de frmConnect.frm")
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    Dim ServerEndpointString As String
    EngineRun = False
    
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me, , False
    
    version.Caption = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
    
    'ServerEndpointString = Replace(ServersLst(1).Ip, ".comunidadargentum.com", "", 1, 10, vbTextCompare)
    'ServerEndpointString = Replace(ServerEndpointString, ".alkononline.com.ar", "", 1, 10, vbTextCompare)
    
    'version = version & " - " & ServerEndpointString
    

    Me.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "VentanaConectar.jpg")
    
    Call LoadButtons

    Call CheckLicenseAgreement
    
    IPTxt.text = IPdelServidor
    PortTxt.text = PuertoDelServidor
    
    Call modCustomCursors.SetFormCursorDefault(Me)

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmConnect.frm")
End Sub

Private Sub CheckLicenseAgreement()
    'Recordatorio para cumplir la licencia, por si borrás el Boton sin leer el code...
On Error GoTo ErrHandler
  
    Dim I As Long
    
    For I = 0 To Me.Controls.Count - 1
        If Me.Controls(I).Name = "imgCodigoFuente" Then
            Exit For
        End If
    Next I
    
    If I = Me.Controls.Count Then
        MsgBox "No debe eliminarse la posibilidad de bajar el código de su servidor. Caso contrario estarían violando la licencia Affero GPL y con ella derechos de autor, incurriendo de esta forma en un delito punible por ley." & vbCrLf & vbCrLf & vbCrLf & _
                "Argentum Online es libre, es de todos. Mantengamoslo así. Si tanto te gusta el juego y querés los cambios que hacemos nosotros, compartí los tuyos. Es un cambio justo. Si no estás de acuerdo, no uses nuestro código, pues nadie te obliga o bien utiliza una versión anterior a la 0.12.0.", vbCritical Or vbApplicationModal
    End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CheckLicenseAgreement de frmConnect.frm")
End Sub

Private Sub LoadButtons()
On Error GoTo ErrHandler
  
    Dim GrhPath As String
    
    GrhPath = DirInterfaces & SELECTED_UI
    
    Set cBotonCrearCuenta = New clsGraphicalButton
    Set cBotonRecuperarPass = New clsGraphicalButton
    Set cButtonWebsite = New clsGraphicalButton
    Set cBotonReglamento = New clsGraphicalButton
    Set cBotonCodigoFuente = New clsGraphicalButton
    Set cBotonBorrarPj = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    Set cBotonLeerMas = New clsGraphicalButton
    Set cButtonDiscord = New clsGraphicalButton
    Set cBotonConectarse = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
                                      
    Call cBotonCrearCuenta.Initialize(imgCrearCuenta, GrhPath & "BotonCrearCuenta.jpg", _
                                    GrhPath & "BotonCrearCuenta.jpg", _
                                    GrhPath & "BotonCrearCuenta.jpg", Me)
                           
                                    
    Call cBotonSalir.Initialize(imgSalir, GrhPath & "BotonSalir.jpg", _
                                    GrhPath & "BotonSalir.jpg", _
                                    GrhPath & "BotonSalir.jpg", Me)

                                    
    Call cBotonConectarse.Initialize(imgConectarse, GrhPath & "BotonIngresar.jpg", _
                                    GrhPath & "BotonIngresar.jpg", _
                                    GrhPath & "BotonIngresar.jpg", Me)
                                    
    Call cButtonDiscord.Initialize(imgDiscord, GrhPath & "BotonDiscord.jpg", _
                                    GrhPath & "BotonDiscord.jpg", _
                                    GrhPath & "BotonDiscord.jpg", Me)
                                    
                                    
    Call cButtonWebsite.Initialize(imgWebsiteLink, "", "", "", Me)

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadButtons de frmConnect.frm")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Public Sub CheckServers()
On Error GoTo ErrHandler
  
    If ServersRecibidos Then
        If Not IsIp(IPTxt) And CurServer <> 0 Then
            If MsgBox("Atencion, está intentando conectarse a un servidor no oficial, NoLand Studios no se hace responsable de los posibles problemas que estos servidores presenten. ¿Desea continuar?", vbYesNo) = vbNo Then
                If CurServer <> 0 Then
                    IPTxt = ServersLst(CurServer).Ip
                    PortTxt = ServersLst(CurServer).Puerto
                Else
                    IPTxt = IPdelServidor
                    PortTxt = PuertoDelServidor
                End If
                Exit Sub
            End If
        End If
    End If
    IPdelServidor = IPTxt
    PuertoDelServidor = PortTxt
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CheckServers de frmConnect.frm")
End Sub


Private Sub imgCodigoFuente_Click()
'***********************************
'IMPORTANTE!
'
'No debe eliminarse la posibilidad de bajar el código de sus servidor de esta forma.
'Caso contrario estarían violando la licencia Affero GPL y con ella derechos de autor,
'incurriendo de esta forma en un delito punible por ley.
'
'Argentum Online es libre, es de todos. Mantengamoslo así. Si tanto te gusta el juego y querés los
'cambios que hacemos nosotros, compartí los tuyos. Es un cambio justo. Si no estás de acuerdo,
'no uses nuestro código, pues nadie te obliga o bien utiliza una versión anterior a la 0.12.0.
'***********************************
On Error GoTo ErrHandler
  
    Call ShellExecute(0, "Open", "https://github.com/argentumonline", "", App.path, SW_SHOWNORMAL)

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgCodigoFuente_Click de frmConnect.frm")
End Sub

Public Sub imgConectarse_Click()

On Error GoTo ErrHandler

      If Not ValidInput(txtNombre.text, txtPasswd.text) Then
        Call frmMessageBox.ShowMessage("Ingrese los datos de autenticación de su cuenta.")
        Exit Sub
    End If
    

    Dim I As Integer
    ' Clean the char slots before login in.
    For I = 1 To ACCPJS
        modAccount.Acc_Data.Acc_Char(I) = EMPTY_CHAR_DATA
    Next I
    
    If Not MainTimer.Check(TimersIndex.Action) Then Exit Sub
        
    GameConfig.Extras.Name = txtNombre.text
        
    Call modAccount.Set_Acc_Data_To_Login
    Call modAccount.Prepare_And_Connect(E_MODO.AccountLogin)

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgConectarse_Click de frmConnect.frm")
End Sub

Public Sub imgCrearCuenta_Click()
    frmAccountCreate.Show , Me
End Sub
  
Private Sub imgDiscord_Click()
    Call Mod_General.OpenDiscordLink
End Sub

Private Sub imgManual_Click()
On Error GoTo ErrHandler
    Call ShellExecute(0, "Open", "https://manual.alkononline.com.ar", "", App.path, SW_SHOWNORMAL)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgManual_Click de frmConnect.frm")
End Sub

Private Sub imgRecuperar_Click()

    Call ShellExecute(0, "Open", ServersLst(CurServer).PanelPassRecoveryUrl, "", App.path, SW_SHOWNORMAL)
    
    'frmAccountRecover.Show , frmConnect

End Sub

Private Sub imgReglamento_Click()
On Error GoTo ErrHandler
    Call ShellExecute(0, "Open", "https://reglamento.alkononline.com.ar", "", App.path, SW_SHOWNORMAL)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgReglamento_Click de frmConnect.frm")
End Sub

Private Sub imgSalir_Click()
    prgRun = False
End Sub

Private Sub imgServArgentina_Click()
On Error GoTo ErrHandler
    Call Engine_Audio.PlayInterface(SND_CLICK)
    IPTxt.text = IPdelServidor

    PortTxt.text = PuertoDelServidor
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgServArgentina_Click de frmConnect.frm")
End Sub

Private Sub imgWebsiteLink_Click()
On Error GoTo ErrHandler
    Call ShellExecute(0, "Open", "https://www.alkononline.com.ar", "", App.path, SW_SHOWNORMAL)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgWebsiteLink_Click de frmConnect.frm")
End Sub

Private Sub txtPasswd_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrHandler
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        imgConectarse_Click
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub txtPasswd_KeyDown de frmConnect.frm")
End Sub
