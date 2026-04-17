VERSION 5.00
Begin VB.Form frmConfigMsg 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3720
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgChckInfo 
      Height          =   195
      Left            =   2090
      Top             =   1710
      Width           =   195
   End
   Begin VB.Image imgChckTrabajo 
      Height          =   195
      Left            =   2090
      Top             =   1370
      Width           =   195
   End
   Begin VB.Image imgChckCombate 
      Height          =   195
      Left            =   400
      Top             =   2080
      Width           =   195
   End
   Begin VB.Image imgChckParty 
      Height          =   195
      Left            =   400
      Top             =   1710
      Width           =   195
   End
   Begin VB.Image imgChckClan 
      Height          =   195
      Left            =   400
      Top             =   1370
      Width           =   195
   End
   Begin VB.Image imgSalir 
      Height          =   375
      Left            =   470
      Top             =   2460
      Width           =   2775
   End
   Begin VB.Image imgGuardar 
      Height          =   375
      Left            =   470
      Top             =   2930
      Width           =   2775
   End
End
Attribute VB_Name = "frmConfigMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private cBotonGuardar As clsGraphicalButton
Private cBotonSalir As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private picCheckBox As Picture

Private clanEnabled As Boolean
Private partyEnabled As Boolean
Private combateEnabled As Boolean
Private trabajoEnabled As Boolean
Private infoEnabled As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CerrarVentana
End Sub


Private Sub CerrarVentana()
On Error GoTo ErrHandler
  
    LoadCustomConsole
    Unload Me
    If frmOpciones.Visible Then
        frmOpciones.SetFocus
    Else
        If frmMain.Visible Then
            frmMain.SetFocus
        End If
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CerrarVentana de frmConfigMsg.frm")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub Form_Load()

On Error GoTo ErrHandler
  
    Me.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "VentanaConfigurarConsola.jpg")
    Call LoadButtons
    
    If Not FileExist(App.path & CustomPath, vbArchive) Then
        Call WriteVar(App.path & CustomPath, "CONFIG", "Clan", 0)
        Call WriteVar(App.path & CustomPath, "CONFIG", "Party", 0)
        Call WriteVar(App.path & CustomPath, "CONFIG", "Combate", 0)
        Call WriteVar(App.path & CustomPath, "CONFIG", "Trabajo", 0)
        Call WriteVar(App.path & CustomPath, "CONFIG", "Info", 0)
    End If
    
    
    clanEnabled = Val(GetVar(App.path & CustomPath, "CONFIG", "Clan"))
    partyEnabled = Val(GetVar(App.path & CustomPath, "CONFIG", "Party"))
    combateEnabled = Val(GetVar(App.path & CustomPath, "CONFIG", "Combate"))
    trabajoEnabled = Val(GetVar(App.path & CustomPath, "CONFIG", "Trabajo"))
    infoEnabled = Val(GetVar(App.path & CustomPath, "CONFIG", "Info"))

    If clanEnabled Then imgChckClan.Picture = picCheckBox
    If partyEnabled Then imgChckParty.Picture = picCheckBox
    If combateEnabled Then imgChckCombate.Picture = picCheckBox
    If trabajoEnabled Then imgChckTrabajo.Picture = picCheckBox
    If infoEnabled Then imgChckInfo.Picture = picCheckBox
    
    Call modCustomCursors.SetFormCursorDefault(Me)

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmConfigMsg.frm")
End Sub

Private Sub LoadButtons()
On Error GoTo ErrHandler
  

    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI
    
    Set cBotonGuardar = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton

    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonGuardar.Initialize(imgGuardar, GrhPath & "BotonConfigurarConsolaGuardar.jpg", _
                                    GrhPath & "BotonConfigurarConsolaGuardarRollover.jpg", _
                                    GrhPath & "BotonConfigurarConsolaGuardarClick.jpg", Me)
                                    
    Call cBotonSalir.Initialize(imgSalir, GrhPath & "BotonConfigurarConsolaSalir.jpg", _
                                    GrhPath & "BotonConfigurarConsolaSalirRollover.jpg", _
                                    GrhPath & "BotonConfigurarConsolaSalirClick.jpg", Me)
                                                             
    Set picCheckBox = LoadPicture(GrhPath & "BotonCheckbox.jpg")
    
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadButtons de frmConfigMsg.frm")
End Sub

Private Sub imgChckClan_Click()
On Error GoTo ErrHandler
  
    clanEnabled = Not clanEnabled
    
    If clanEnabled Then
        imgChckClan.Picture = picCheckBox
    Else
        Set imgChckClan.Picture = Nothing
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgChckClan_Click de frmConfigMsg.frm")
End Sub

Private Sub imgChckCombate_Click()
On Error GoTo ErrHandler
  
    combateEnabled = Not combateEnabled
    
    If combateEnabled Then
        imgChckCombate.Picture = picCheckBox
    Else
        Set imgChckCombate.Picture = Nothing
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgChckCombate_Click de frmConfigMsg.frm")
End Sub

Private Sub imgChckInfo_Click()
On Error GoTo ErrHandler
  
    infoEnabled = Not infoEnabled
    
    If infoEnabled Then
        imgChckInfo.Picture = picCheckBox
    Else
        Set imgChckInfo.Picture = Nothing
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgChckInfo_Click de frmConfigMsg.frm")
End Sub

Private Sub imgChckParty_Click()
On Error GoTo ErrHandler
  
    partyEnabled = Not partyEnabled
    
    If partyEnabled Then
        imgChckParty.Picture = picCheckBox
    Else
        Set imgChckParty.Picture = Nothing
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgChckParty_Click de frmConfigMsg.frm")
End Sub

Private Sub imgChckTrabajo_Click()
On Error GoTo ErrHandler
  
    trabajoEnabled = Not trabajoEnabled
    
    If trabajoEnabled Then
        imgChckTrabajo.Picture = picCheckBox
    Else
        Set imgChckTrabajo.Picture = Nothing
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgChckTrabajo_Click de frmConfigMsg.frm")
End Sub

Private Sub imgGuardar_Click()
        
On Error GoTo ErrHandler
  

    Call WriteVar(App.path & CustomPath, "CONFIG", "Clan", CInt(clanEnabled) * -1)
    Call WriteVar(App.path & CustomPath, "CONFIG", "Party", CInt(partyEnabled) * -1)
    Call WriteVar(App.path & CustomPath, "CONFIG", "Combate", CInt(combateEnabled) * -1)
    Call WriteVar(App.path & CustomPath, "CONFIG", "Trabajo", CInt(trabajoEnabled) * -1)
    Call WriteVar(App.path & CustomPath, "CONFIG", "Info", CInt(infoEnabled) * -1)
    
    Call CerrarVentana

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgGuardar_Click de frmConfigMsg.frm")
End Sub

Private Sub imgSalir_Click()
    Call CerrarVentana
End Sub

