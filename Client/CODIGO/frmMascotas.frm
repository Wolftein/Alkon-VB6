VERSION 5.00
Begin VB.Form frmMascotas 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Mascotas"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picMascota 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1190
      Index           =   0
      Left            =   640
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   2
      Top             =   628
      Width           =   1190
   End
   Begin VB.PictureBox picMascota 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1190
      Index           =   1
      Left            =   2320
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   1
      Top             =   628
      Width           =   1190
   End
   Begin VB.PictureBox picMascota 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1190
      Index           =   2
      Left            =   4000
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   0
      Top             =   628
      Width           =   1190
   End
   Begin VB.Image imgBox 
      Height          =   1500
      Index           =   3
      Left            =   3840
      Top             =   480
      Width           =   1500
   End
   Begin VB.Image imgBox 
      Height          =   1500
      Index           =   2
      Left            =   2160
      Top             =   480
      Width           =   1500
   End
   Begin VB.Image imgBox 
      Height          =   1500
      Index           =   1
      Left            =   480
      Top             =   480
      Width           =   1500
   End
   Begin VB.Image imgSeleccionar 
      Height          =   660
      Left            =   3000
      Top             =   2400
      Width           =   1425
   End
   Begin VB.Image imgCerrar 
      Height          =   255
      Left            =   5385
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgLiberar 
      Height          =   660
      Left            =   1200
      Top             =   2400
      Width           =   1425
   End
End
Attribute VB_Name = "frmMascotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Private cBotonCerrar As clsGraphicalButton
Private cBotonLiberar As clsGraphicalButton
Private cBotonSeleccionar As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Devices(0 To 2) As Long

Private Sub Form_KeyDown2(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CerrarVentana
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CerrarVentana
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
  
    'Set the form manager to allow the form to be moved
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    Call LoadButtons
    Call SelectBox(PetSelectedIndex)
    
    Dim I As Long
    For I = 0 To 2
        Devices(I) = Aurora_Graphics.CreatePassFromDisplay(picMascota(I).hwnd, picMascota(I).ScaleWidth, picMascota(I).ScaleHeight)
        
        Call Invalidate(picMascota(I).hwnd)
    Next I
    
    Call modCustomCursors.SetFormCursorDefault(Me)

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmMascotas.frm")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Dim I As Long
    For I = 0 To 2
        Call Aurora_Graphics.DeletePass(Devices(I))
    Next I
    
End Sub

Private Sub LoadButtons()
On Error GoTo ErrHandler
  
    Dim GrhPath As String
    
    GrhPath = DirInterfaces & SELECTED_UI

    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonLiberar = New clsGraphicalButton
    Set cBotonSeleccionar = New clsGraphicalButton
    Set LastButtonPressed = New clsGraphicalButton
    
    Me.Picture = LoadPicture(GrhPath & "VentanaMascotas.jpg")

    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonMascotaSalir.jpg", GrhPath & "BotonMascotaSalirRollover.jpg", GrhPath & "BotonMascotaSalirClick.jpg", Me)
    Call cBotonLiberar.Initialize(imgLiberar, GrhPath & "BotonMascotasLiberar.jpg", GrhPath & "BotonMascotasLiberarRollover.jpg", GrhPath & "BotonMascotasLiberarClick.jpg", Me, GrhPath & "BotonMascotasLiberarDisabled.jpg")
    Call cBotonSeleccionar.Initialize(imgSeleccionar, GrhPath & "BotonMascotasSeleccionar.jpg", GrhPath & "BotonMascotasSeleccionarRollover.jpg", GrhPath & "BotonMascotasSeleccionarClick.jpg", Me, GrhPath & "BotonMascotasSeleccionarDisabled.jpg")
      
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadButtons de frmMascotas.frm")
End Sub

Private Sub imgCerrar_Click()
    CerrarVentana
End Sub

Private Sub CerrarVentana()
On Error GoTo ErrHandler
  
    Unload Me
    If frmMain.Visible Then frmMain.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CerrarVentana de frmCraft.frm")
End Sub

Private Sub imgLiberar_Click()
    Call WriteReleasePet(True, PetSelectedIndex)
End Sub

Private Sub imgSeleccionar_Click()
    Call WriteSelectPet(PetSelectedIndex)
End Sub

Private Sub SelectBox(ByVal element As Byte)
    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI

    Dim I As Integer
    For I = 1 To 3
        If I = element Then
            imgBox(I).Picture = LoadPicture(GrhPath & "CajaMascotaClick.jpg")
        Else
            imgBox(I).Picture = LoadPicture(GrhPath & "CajaMascota.jpg")
        End If
    Next I
End Sub

Public Sub RefreshBoxes()
    Call SelectBox(PetSelectedIndex)
End Sub

Private Sub picMascota_Click(Index As Integer)
    PetSelectedIndex = Index + 1
    Call SelectBox(PetSelectedIndex)
End Sub

Private Sub picMascota_Paint(Index As Integer)
    Dim Grh As GrhData
    Dim X As Integer
    Dim Y As Integer
                    
    Dim PetIndex As Long
    PetIndex = Index + 1
        
    Call UIBegin(Devices(Index), picMascota(Index).ScaleWidth, picMascota(Index).ScaleHeight, 0)
    
    If PetIndex <= PetListQty Then
        If PetList(PetIndex) <> 0 Then
            Grh = GrhData(BodyData(PetList(PetIndex)).Walk(SOUTH).GrhIndex)
    
            X = (picMascota(Index).ScaleWidth / 2) - (Grh.pixelWidth / 2)
            Y = (picMascota(Index).ScaleHeight / 2) - (Grh.pixelHeight / 2)
    
            Call DrawGrhIndex(Grh.Frames(1), X, Y, 0#, False)
        End If
    End If
    
    Call UIEnd

End Sub

