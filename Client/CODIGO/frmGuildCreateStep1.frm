VERSION 5.00
Begin VB.Form frmGuildCreateStep1 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "z"
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image ImgSiguiente 
      Height          =   570
      Left            =   3260
      Top             =   4350
      Width           =   1290
   End
   Begin VB.Image ImgCancelar 
      Height          =   570
      Left            =   1430
      Top             =   4350
      Width           =   1290
   End
   Begin VB.Label LblUserAlign 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LblUserAlign"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   1880
      Width           =   1095
   End
   Begin VB.Label LblAlignDescrip 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGuildCreateStep1.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1790
      Left            =   750
      TabIndex        =   1
      Top             =   2350
      Width           =   4370
   End
   Begin VB.Label LblAlign 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "La alineación del clan sera dada por la alineacion del lider al momento de ser creada, esta no podra se cambiada luego."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   5295
   End
End
Attribute VB_Name = "frmGuildCreateStep1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Align As eGuildAlignment

Private clsFormulario As clsFormMovementManager
Private cButtonCancel As clsGraphicalButton
Private cButtonNext As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CloseWindow
End Sub

Private Sub Form_Load()

    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me, , False

    Call LoadControls
    Call LoadControlData
    
    Call modCustomCursors.SetFormCursorDefault(Me)
End Sub

Public Sub SetAlignment(ByVal Alignment As eGuildAlignment)
   Align = Alignment
End Sub

Private Sub imgCancelar_Click()
    Unload frmGuildCreateStep1
    Unload frmGuildCreateStep2
End Sub

Private Sub ImgSiguiente_Click()
    frmGuildCreateStep2.Show , frmMain
    frmGuildCreateStep1.Visible = False
End Sub

Private Sub LoadControls()
    
    Set cButtonCancel = New clsGraphicalButton
    Set cButtonNext = New clsGraphicalButton

    
    Set LastButtonPressed = New clsGraphicalButton
    
    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI
    
    Me.Picture = LoadPicture(GrhPath & "VentanaGuildCreationStep1.jpg")
    
    Call cButtonCancel.Initialize(ImgCancelar, GrhPath & "BotonGuildCreationCancelar.jpg", _
                                    GrhPath & "BotonGuildCreationCancelar.jpg", _
                                    GrhPath & "BotonGuildCreationCancelar.jpg", Me)
                                    'GrhPath & "BotonGuildCreationCancelar.jpg", _
                                    'GrhPath & "BotonGuildCreationCancelar.jpg", Me)
                                    
    Call cButtonNext.Initialize(ImgSiguiente, GrhPath & "BotonGuildCreationSiguiente.jpg", _
                                    GrhPath & "BotonGuildCreationSiguiente.jpg", _
                                    GrhPath & "BotonGuildCreationSiguiente.jpg", Me)
                                    'GrhPath & "BotonGuildCreationSiguiente.jpg", _
                                    'GrhPath & "BotonGuildCreationSiguiente.jpg", Me)

    
End Sub

Private Sub LoadControlData()

    Dim GuildToCreate As tGuildInfo
    Dim RealText As String, NeutralText As String, EvilText, AlignText As String
    
    AlignText = "La alineación del clan sera dada por la alineacion del lider al momento de ser creado. Ésta no podra se cambiada luego."
    
    RealText = "En los clanes de alineación Real solo se admiten miembros del Ejercito Real de Banderbill. Lucharan contra todos los criminales de estas tierras siempre junto al Rey y bajo las órdenes del Consejo de Banderbill. Los personajes que no sean miembros del ejército no podrán pertenecer a un clan Real."
    NeutralText = "Los clanes neutrales son aquellos que no tienen ninguna alineación manifiesta. Solo se aceptarán personajes neutrales que no pertenezcan a ninguna facción."
    EvilText = "Estos Clanes estarán formados únicamente por miembros de la Legión Oscura y bajo las órdenes del demonio y sus súbditos. Los clanes que le juren lealtad no podrán tener miembros neutrales o de alineación Real."
    
    'set default data
    GuildCreation = GuildToCreate
    GuildCreation.Alignment = Align
    
    LblAlign.Caption = AlignText
    LblUserAlign.Caption = GetNameOfAlignment(GuildCreation.Alignment)
    
    Select Case Align
        Case IsReal
            LblAlignDescrip = RealText
        Case IsNeutral
            LblAlignDescrip = NeutralText
        Case IsEvil
            LblAlignDescrip = EvilText
    End Select
    
End Sub

Private Sub CloseWindow()
On Error GoTo ErrHandler
  
    Unload Me
    If frmMain.Visible Then frmMain.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmGuildCreateStep1.frm")
End Sub
