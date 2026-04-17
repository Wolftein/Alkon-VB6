VERSION 5.00
Begin VB.Form frmMapa 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNpcs 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1165
      Left            =   9060
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4640
      Width           =   2635
   End
   Begin VB.TextBox txtEntradas 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1165
      Left            =   9060
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   6495
      Width           =   2635
   End
   Begin VB.PictureBox picMapa 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   8640
      Left            =   150
      ScaleHeight     =   576
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   576
      TabIndex        =   0
      Top             =   180
      Width           =   8640
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9960
      TabIndex        =   7
      Top             =   2370
      Width           =   1695
   End
   Begin VB.Label lblRegion 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9960
      TabIndex        =   6
      Top             =   2920
      Width           =   1695
   End
   Begin VB.Image imgMostrar 
      Height          =   420
      Left            =   9480
      Top             =   360
      Width           =   420
   End
   Begin VB.Image imgCerrar 
      Height          =   195
      Left            =   11640
      MousePointer    =   99  'Custom
      Top             =   150
      Width           =   195
   End
   Begin VB.Image imgMapaDungeon 
      Height          =   330
      Left            =   9120
      Top             =   1170
      Width           =   2610
   End
   Begin VB.Image imgMapaGeneral 
      Height          =   330
      Left            =   9120
      Top             =   840
      Width           =   2610
   End
   Begin VB.Label lblNivel 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10950
      TabIndex        =   5
      Top             =   3455
      Width           =   615
   End
   Begin VB.Label lblMapa 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9960
      TabIndex        =   4
      Top             =   2060
      Width           =   615
   End
   Begin VB.Label lblZona 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9960
      TabIndex        =   3
      Top             =   3780
      Width           =   1575
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsFormulario As clsFormMovementManager
Private BotonGeneral As clsGraphicalButton
Private BotonDungeon As clsGraphicalButton
Private BotonCerrar As clsGraphicalButton

Private BotonTic As Picture
Private BotonTac As Picture

Public OnFocus As Byte

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Activate()
    OnFocus = 1
End Sub

Private Sub Form_Deactivate()
    OnFocus = 0
End Sub

Private Sub Form_GotFocus()
    OnFocus = 1
End Sub

Private Sub Form_LostFocus()
    OnFocus = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub Form_Load()
    
On Error GoTo ErrHandler
  
Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Set MapaGrafico = New clsGraphicalMap
    
    Call InicializarMapa(0)
    MapaGrafico.SetVerNumeros = 1
    
    OnFocus = 1

    Call LoadButtons
    
    Call BotonGeneral.EnableButton(False)
    
    Call modCustomCursors.SetFormCursorDefault(Me)

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmMapa.frm")
End Sub

Private Sub LoadButtons()
On Error GoTo ErrHandler
  
    Dim GrhPath As String
    
    GrhPath = DirInterfaces & SELECTED_UI
    
    Set BotonGeneral = New clsGraphicalButton
    Set BotonDungeon = New clsGraphicalButton
    Set BotonCerrar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton

    Call BotonGeneral.Initialize(imgMapaGeneral, GrhPath & "BotonMapaMundo.jpg", _
                                    GrhPath & "BotonMapaMundoRollover.jpg", _
                                    GrhPath & "BotonMapaMundoClick.jpg", Me, GrhPath & "BotonMapaMundoClick.jpg")
                                    
    Call BotonDungeon.Initialize(imgMapaDungeon, GrhPath & "BotonMapaDungeons.jpg", _
                                    GrhPath & "BotonMapaDungeonsRollover.jpg", _
                                    GrhPath & "BotonMapaDungeonsClick.jpg", Me, GrhPath & "BotonMapaDungeonsClick.jpg")

    Call BotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCruzSalir.jpg", _
                                    GrhPath & "BotonCruzSalirRollover.jpg", _
                                    GrhPath & "BotonCruzSalirClick.jpg", Me)
    
    
    Set BotonTic = LoadPicture(GrhPath & "BotonDueloAmigosTic.jpg")
    Set BotonTac = LoadPicture(GrhPath & "BotonDueloAmigosTac.jpg")
    
    Me.Picture = LoadPicture(GrhPath & "VentanaMapa.jpg")
    imgMostrar.Picture = BotonTac
    Me.Icon = frmMain.Icon
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadButtons de frmMapa.frm")
End Sub

Private Sub imgMapaGeneral_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgMapaDungeon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgMapaDungeon_Click()
On Error GoTo ErrHandler
  
    If MapaGrafico.GetMapType = 1 Then Exit Sub
    MapaGrafico.SetMapType = 1
    Call InicializarMapa(1)

    Call BotonGeneral.EnableButton(True)
    Call BotonDungeon.EnableButton(False)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgMapaDungeon_Click de frmMapa.frm")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
        If frmMain.Visible Then frmMain.SetFocus
    End If
End Sub

Private Sub picMapa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgMostrar_Click()
On Error GoTo ErrHandler
  
    If MapaGrafico.GetVerNumeros = 0 Then
        imgMostrar.Picture = BotonTac
        MapaGrafico.SetVerNumeros = 1
    Else
        imgMostrar.Picture = BotonTic
        MapaGrafico.SetVerNumeros = 0
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgMostrar_Click de frmMapa.frm")
End Sub

Private Sub imgMapaGeneral_Click()
On Error GoTo ErrHandler
  
    If MapaGrafico.GetMapType = 0 Then Exit Sub
    MapaGrafico.SetMapType = 0
    Call InicializarMapa(0)

    Call BotonGeneral.EnableButton(False)
    Call BotonDungeon.EnableButton(True)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgMapaGeneral_Click de frmMapa.frm")
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrHandler
  
    Select Case KeyCode
        Case vbKeyUp
            Call MapaGrafico.DoScroll(1)
        Case vbKeyRight
            Call MapaGrafico.DoScroll(2)
        Case vbKeyDown
            Call MapaGrafico.DoScroll(3)
        Case vbKeyLeft
            Call MapaGrafico.DoScroll(4)
    End Select
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_KeyUp de frmMapa.frm")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandler
  
    Set MapaGrafico = Nothing
    OnFocus = 0
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Unload de frmMapa.frm")
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub
