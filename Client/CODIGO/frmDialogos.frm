VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmDialogos 
   BorderStyle     =   0  'None
   Caption         =   "Dialogos"
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Slider b 
      Height          =   255
      Left            =   2990
      TabIndex        =   2
      Top             =   1620
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      LargeChange     =   25
      Max             =   255
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider g 
      Height          =   255
      Left            =   2990
      TabIndex        =   1
      Top             =   1260
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      LargeChange     =   25
      Max             =   255
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider r 
      Height          =   255
      Left            =   2990
      TabIndex        =   0
      Top             =   900
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      LargeChange     =   25
      Max             =   255
      TickStyle       =   3
   End
   Begin VB.Image imgSalir 
      Height          =   195
      Left            =   5550
      Top             =   150
      Width           =   195
   End
   Begin VB.Image imgOption 
      Height          =   195
      Index           =   6
      Left            =   460
      MousePointer    =   99  'Custom
      Top             =   2380
      Width           =   195
   End
   Begin VB.Image imgOption 
      Height          =   195
      Index           =   5
      Left            =   460
      MousePointer    =   99  'Custom
      Top             =   2160
      Width           =   195
   End
   Begin VB.Image imgOption 
      Height          =   195
      Index           =   4
      Left            =   460
      MousePointer    =   99  'Custom
      Top             =   1930
      Width           =   195
   End
   Begin VB.Image imgOption 
      Height          =   195
      Index           =   3
      Left            =   460
      MousePointer    =   99  'Custom
      Top             =   1710
      Width           =   195
   End
   Begin VB.Image imgOption 
      Height          =   195
      Index           =   2
      Left            =   460
      MousePointer    =   99  'Custom
      Top             =   1480
      Width           =   195
   End
   Begin VB.Image imgOption 
      Height          =   195
      Index           =   1
      Left            =   460
      MousePointer    =   99  'Custom
      Top             =   1250
      Width           =   195
   End
   Begin VB.Image imgTodosPorDefecto 
      Height          =   375
      Left            =   250
      Top             =   2940
      Width           =   1935
   End
   Begin VB.Image imgPorDefecto 
      Height          =   375
      Left            =   3320
      Top             =   1980
      Width           =   1935
   End
   Begin VB.Image imgAceptar 
      Height          =   375
      Left            =   3700
      Top             =   2940
      Width           =   1935
   End
   Begin VB.Label lblPrueba 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   820
      Left            =   2360
      TabIndex        =   3
      Top             =   2520
      Width           =   1190
   End
End
Attribute VB_Name = "frmDialogos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private picOptionInactive As Picture
Private picOptionActive As Picture

Private cBotonAceptar As clsGraphicalButton
Private cBotonAplicar As clsGraphicalButton
Private cBotonPorDefecto As clsGraphicalButton
Private cBotonTodosPorDefecto As clsGraphicalButton
Private cBotonSalir As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private indice As Byte
Private r_selected As Byte
Private g_selected As Byte
Private b_selected As Byte

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CloseWindow
End Sub

'***************************************************
'Author: Juan Dalmasso (CHOTS)
'Last Modify Date: 12/06/2011
'Allow the character to modify the color of each dialog in the game
'***************************************************

Private Sub Form_Load()

    ' Handles Form movement (drag and drop).
On Error GoTo ErrHandler
  
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "VentanaDialogos.JPG")
    
    indice = 1
    
    r.value = ColoresDialogos(indice).r
    r_selected = r.value
    g.value = ColoresDialogos(indice).g
    g_selected = g.value
    b.value = ColoresDialogos(indice).b
    b_selected = b.value
    
    Call LoadButtons
    
    Call modCustomCursors.SetFormCursorDefault(Me)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmDialogos.frm")
End Sub

Private Sub LoadButtons()
On Error GoTo ErrHandler
  
    Dim GrhPath As String
    
    GrhPath = DirInterfaces & SELECTED_UI

    Set cBotonAceptar = New clsGraphicalButton
    Set cBotonAplicar = New clsGraphicalButton
    Set cBotonPorDefecto = New clsGraphicalButton
    Set cBotonTodosPorDefecto = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    
    
    Set LastButtonPressed = New clsGraphicalButton
    
    
    Call cBotonAceptar.Initialize(imgAceptar, GrhPath & "BotonDialogosAplicar.JPG", _
                                    GrhPath & "BotonDialogosAplicarRollover.JPG", _
                                    GrhPath & "BotonDialogosAplicarClick.JPG", Me)

    Call cBotonPorDefecto.Initialize(imgPorDefecto, GrhPath & "BotonDialogosDefecto.JPG", _
                                    GrhPath & "BotonDialogosDefectoRollover.JPG", _
                                    GrhPath & "BotonDialogosDefectoClick.JPG", Me)
                                    
    Call cBotonTodosPorDefecto.Initialize(imgTodosPorDefecto, GrhPath & "BotonDialogosTodoDefecto.JPG", _
                                    GrhPath & "BotonDialogosTodoDefectoRollover.JPG", _
                                    GrhPath & "BotonDialogosTodoDefectoClick.JPG", Me)
                                    
    Call cBotonSalir.Initialize(imgSalir, GrhPath & "BotonDialogosSalir.JPG", _
                                    GrhPath & "BotonDialogosSalirRollover.JPG", _
                                    GrhPath & "BotonDialogosSalirClick.JPG", Me)
                                    
    Set picOptionInactive = Nothing 'LoadPicture(GrhPath & "BotonDialogosSalirDisabled.jpg")
    Set picOptionActive = LoadPicture(GrhPath & "BotonDialogosSalir.jpg")
    
    ' Seleccionado el primero por default
    Set imgOption(indice).Picture = picOptionActive
    imgOption(indice).MouseIcon = picMouseIcon
    
    Dim lOption As Long
    For lOption = 2 To MAXCOLORESDIALOGOS
        Set imgOption(lOption).Picture = picOptionInactive
        imgOption(lOption).MouseIcon = picMouseIcon
    Next lOption
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadButtons de frmDialogos.frm")
End Sub

Private Sub b_Change()
    lblPrueba.BackColor = RGB(r.value, g.value, b.value)
On Error GoTo ErrHandler
  
    b_selected = b.value
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub b_Change de frmDialogos.frm")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub Image1_Click()
End Sub

Private Sub imgAceptar_Click()
    'hace un "actualizar" antes de cerrarlo
On Error GoTo ErrHandler
  
    ColoresDialogos(indice).r = r_selected
    ColoresDialogos(indice).g = g_selected
    ColoresDialogos(indice).b = b_selected
    
    'los graba en el dialogos.dat
    Call GrabarColores
    
    Unload frmDialogos
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgAceptar_Click de frmDialogos.frm")
End Sub

Private Sub imgAplicar_Click()
    'ColoresDialogos(indice).r = r.value
    'ColoresDialogos(indice).g = g.value
    'ColoresDialogos(indice).b = b.value
End Sub

Private Sub imgOption_Click(Index As Integer)
    
On Error GoTo ErrHandler
  
    If Index <> indice Then
        Set imgOption(indice).Picture = picOptionInactive
        Set imgOption(Index).Picture = picOptionActive
    End If
    
    r.value = ColoresDialogos(Index).r
    g.value = ColoresDialogos(Index).g
    b.value = ColoresDialogos(Index).b
    indice = Index
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgOption_Click de frmDialogos.frm")
End Sub

Private Sub imgPorDefecto_Click()
    Call PorDefecto(indice)
End Sub

Private Sub imgSalir_Click()
On Error GoTo ErrHandler
  
    Call CargarColores
    Unload frmDialogos
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgSalir_Click de frmDialogos.frm")
End Sub

Private Sub imgTodosPorDefecto_Click()
On Error GoTo ErrHandler
  
    Call TodosPorDefecto
    
    r.value = ColoresDialogos(indice).r
    g.value = ColoresDialogos(indice).g
    b.value = ColoresDialogos(indice).b
    
    lblPrueba.BackColor = RGB(r.value, g.value, b.value)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgTodosPorDefecto_Click de frmDialogos.frm")
End Sub

Private Sub g_Change()
On Error GoTo ErrHandler
  
    lblPrueba.BackColor = RGB(r.value, g.value, b.value)
    g_selected = g.value
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub g_Change de frmDialogos.frm")
End Sub

Private Sub r_Change()
On Error GoTo ErrHandler
  
    lblPrueba.BackColor = RGB(r.value, g.value, b.value)
    r_selected = r.value
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub r_Change de frmDialogos.frm")
End Sub

Private Sub PorDefecto(ByVal I As Byte)
On Error GoTo ErrHandler
  
    Dim archivoC As String
    archivoC = App.path & "\init\DialogosBACKUP.dat"
    
    ColoresDialogos(I).r = CByte(GetVar(archivoC, CStr(I), "R"))
    
    ColoresDialogos(I).g = CByte(GetVar(archivoC, CStr(I), "G"))
   
    ColoresDialogos(I).b = CByte(GetVar(archivoC, CStr(I), "B"))

    
    r.value = ColoresDialogos(I).r
    g.value = ColoresDialogos(I).g
    b.value = ColoresDialogos(I).b
    
    r_selected = ColoresDialogos(I).r
    g_selected = ColoresDialogos(I).g
    b_selected = ColoresDialogos(I).b
    
    lblPrueba.BackColor = RGB(r.value, g.value, b.value)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub PorDefecto de frmDialogos.frm")
End Sub

Private Sub TodosPorDefecto()
On Error GoTo ErrHandler
  

    Dim I As Byte
    
    For I = 1 To MAXCOLORESDIALOGOS
        Call PorDefecto(I)
    Next I

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub TodosPorDefecto de frmDialogos.frm")
End Sub

Private Sub GrabarColores()
On Error GoTo ErrHandler
  
    Dim archivoC As String
    Dim I As Byte
    archivoC = App.path & "\init\Dialogos.dat"
    
    For I = 1 To MAXCOLORESDIALOGOS
        Call WriteVar(archivoC, I, "R", ColoresDialogos(I).r)
        Call WriteVar(archivoC, I, "G", ColoresDialogos(I).g)
        Call WriteVar(archivoC, I, "B", ColoresDialogos(I).b)
    Next I
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GrabarColores de frmDialogos.frm")
End Sub

Private Sub CloseWindow()
On Error GoTo ErrHandler
  
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
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmDialogos.frm")
End Sub
