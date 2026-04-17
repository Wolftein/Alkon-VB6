VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmTournament 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurar Torneo"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5385
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Peleas"
      TabPicture(0)   =   "frmTournament.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Configuración"
      TabPicture(1)   =   "frmTournament.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   6900
         Left            =   0
         TabIndex        =   20
         Top             =   320
         Width           =   5415
         Begin VB.Frame Frame5 
            Caption         =   "Mapas"
            Height          =   975
            Left            =   120
            TabIndex        =   42
            Top             =   5280
            Width           =   5055
            Begin VB.PictureBox picPositions 
               BorderStyle     =   0  'None
               Height          =   735
               Index           =   0
               Left            =   2880
               ScaleHeight     =   735
               ScaleWidth      =   2055
               TabIndex        =   46
               Top             =   120
               Width           =   2055
               Begin VB.TextBox txtPosX 
                  Height          =   285
                  Left            =   360
                  TabIndex        =   49
                  Text            =   "0"
                  Top             =   360
                  Width           =   495
               End
               Begin VB.TextBox txtPosY 
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   47
                  Text            =   "0"
                  Top             =   360
                  Width           =   495
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "X:"
                  Height          =   195
                  Index           =   7
                  Left            =   120
                  TabIndex        =   50
                  Top             =   390
                  Width           =   195
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Y:"
                  Height          =   195
                  Index           =   5
                  Left            =   1080
                  TabIndex        =   48
                  Top             =   390
                  Width           =   195
               End
            End
            Begin VB.TextBox txtMap 
               Height          =   285
               Left            =   2040
               TabIndex        =   44
               Text            =   "0"
               Top             =   480
               Width           =   615
            End
            Begin VB.ComboBox cboMaps 
               Height          =   315
               Left            =   120
               TabIndex        =   43
               Text            =   "Combo1"
               Top             =   480
               Width           =   1695
            End
            Begin VB.PictureBox picPositions 
               BorderStyle     =   0  'None
               Height          =   735
               Index           =   1
               Left            =   2880
               ScaleHeight     =   735
               ScaleWidth      =   2055
               TabIndex        =   51
               Top             =   120
               Width           =   2055
               Begin VB.TextBox txtUser1Y 
                  Height          =   285
                  Left            =   1560
                  TabIndex        =   57
                  Text            =   "0"
                  Top             =   120
                  Width           =   495
               End
               Begin VB.TextBox txtUser1X 
                  Height          =   285
                  Left            =   600
                  TabIndex        =   56
                  Text            =   "0"
                  Top             =   120
                  Width           =   495
               End
               Begin VB.TextBox txtUser2Y 
                  Height          =   285
                  Left            =   1560
                  TabIndex        =   53
                  Text            =   "0"
                  Top             =   480
                  Width           =   495
               End
               Begin VB.TextBox txtUser2X 
                  Height          =   285
                  Left            =   600
                  TabIndex        =   52
                  Text            =   "0"
                  Top             =   480
                  Width           =   495
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  Caption         =   "U2"
                  Height          =   195
                  Left            =   0
                  TabIndex        =   61
                  Top             =   510
                  Width           =   255
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  Caption         =   "U1"
                  Height          =   195
                  Left            =   0
                  TabIndex        =   60
                  Top             =   150
                  Width           =   255
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Y:"
                  Height          =   195
                  Index           =   11
                  Left            =   1320
                  TabIndex        =   59
                  Top             =   150
                  Width           =   195
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "X:"
                  Height          =   195
                  Index           =   10
                  Left            =   360
                  TabIndex        =   58
                  Top             =   150
                  Width           =   195
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Y:"
                  Height          =   195
                  Index           =   9
                  Left            =   1320
                  TabIndex        =   55
                  Top             =   510
                  Width           =   195
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "X:"
                  Height          =   195
                  Index           =   8
                  Left            =   360
                  TabIndex        =   54
                  Top             =   510
                  Width           =   195
               End
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mapa"
               Height          =   195
               Index           =   6
               Left            =   2040
               TabIndex        =   45
               Top             =   240
               Width           =   480
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Items Prohibidos"
            Height          =   1815
            Left            =   120
            TabIndex        =   37
            Top             =   3360
            Width           =   5055
            Begin VB.ListBox lstItemsProhibidos 
               Height          =   1185
               Left            =   120
               Style           =   1  'Checkbox
               TabIndex        =   39
               Top             =   480
               Width           =   2655
            End
            Begin VB.ListBox lstItemsProhibidosSel 
               Height          =   1230
               Left            =   2760
               TabIndex        =   38
               Top             =   480
               Width           =   2175
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Seleccionados"
               Height          =   255
               Index           =   1
               Left            =   2760
               TabIndex        =   41
               Top             =   240
               Width           =   2055
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Selección"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   40
               Top             =   240
               Width           =   2055
            End
         End
         Begin VB.TextBox txtNumRounds 
            Height          =   285
            Left            =   4080
            TabIndex        =   36
            Text            =   "0"
            Top             =   960
            Width           =   975
         End
         Begin VB.ListBox lstClasesPermitidas 
            Height          =   735
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   33
            Top             =   2520
            Width           =   2775
         End
         Begin VB.CommandButton cmdGuardarConfig 
            Caption         =   "Guardar"
            Height          =   375
            Left            =   3120
            TabIndex        =   32
            Top             =   6360
            Width           =   2055
         End
         Begin VB.TextBox txtNroParticipantes 
            Height          =   285
            Left            =   2160
            TabIndex        =   31
            Text            =   "0"
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox txtOroRequerido 
            Height          =   285
            Left            =   2160
            TabIndex        =   29
            Text            =   "0"
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CheckBox chkEjecutarAlMorir 
            Caption         =   "Ejecutar Competidores al morir"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   6480
            Width           =   3015
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nivel Requerido"
            Height          =   1335
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   1815
            Begin VB.TextBox txtMAxNivel 
               Height          =   285
               Left            =   600
               TabIndex        =   26
               Text            =   "0"
               Top             =   840
               Width           =   975
            End
            Begin VB.TextBox txtMinNivel 
               Height          =   285
               Left            =   600
               TabIndex        =   25
               Text            =   "0"
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Max:"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   24
               Top             =   840
               Width           =   855
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Min:"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   23
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.CommandButton cmdConfiguracionActual 
            Caption         =   "Mostrar Configuración actual"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   5175
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rounds"
            Height          =   195
            Index           =   4
            Left            =   4080
            TabIndex        =   35
            Top             =   720
            Width           =   660
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Clases Permitidas"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   2280
            Width           =   2055
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Participantes"
            Height          =   195
            Index           =   3
            Left            =   2160
            TabIndex        =   30
            Top             =   720
            Width           =   1545
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Oro Requerido"
            Height          =   195
            Index           =   2
            Left            =   2160
            TabIndex        =   28
            Top             =   1440
            Width           =   1245
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6900
         Index           =   0
         Left            =   -75000
         TabIndex        =   1
         Top             =   320
         Width           =   5415
         Begin VB.CommandButton cmdActualizarListaParticipantes 
            Caption         =   "Actualizar Lista"
            Height          =   255
            Left            =   3360
            TabIndex        =   19
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton cmdCancelarTorneo 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   3600
            TabIndex        =   18
            Top             =   6360
            Width           =   1575
         End
         Begin VB.Frame Frame1 
            Caption         =   "Pelea"
            Height          =   2535
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Top             =   3480
            Width           =   5055
            Begin VB.OptionButton optArena 
               Caption         =   "5"
               Height          =   255
               Index           =   4
               Left            =   3600
               TabIndex        =   15
               Top             =   1680
               Width           =   495
            End
            Begin VB.OptionButton optArena 
               Caption         =   "4"
               Height          =   255
               Index           =   3
               Left            =   3000
               TabIndex        =   14
               Top             =   1680
               Width           =   495
            End
            Begin VB.OptionButton optArena 
               Caption         =   "3"
               Height          =   255
               Index           =   2
               Left            =   2400
               TabIndex        =   13
               Top             =   1680
               Width           =   495
            End
            Begin VB.OptionButton optArena 
               Caption         =   "2"
               Height          =   255
               Index           =   1
               Left            =   1800
               TabIndex        =   12
               Top             =   1680
               Width           =   495
            End
            Begin VB.OptionButton optArena 
               Caption         =   "1"
               Height          =   255
               Index           =   0
               Left            =   1200
               TabIndex        =   11
               Top             =   1680
               Value           =   -1  'True
               Width           =   495
            End
            Begin VB.CommandButton cmdQuitarParticipante1 
               Caption         =   "X"
               Height          =   375
               Left            =   4560
               TabIndex        =   10
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton cmdQuitarParticipante2 
               Caption         =   "X"
               Height          =   375
               Left            =   4560
               TabIndex        =   9
               Top             =   1080
               Width           =   375
            End
            Begin VB.CommandButton cmdComenzarPelea 
               Caption         =   "Comenzar Pelea"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   8
               Top             =   2040
               Width           =   4815
            End
            Begin VB.TextBox txtParticipante2 
               Height          =   375
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   7
               Top             =   1080
               Width           =   4335
            End
            Begin VB.TextBox txtParticipante1 
               Height          =   375
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   6
               Top             =   240
               Width           =   4335
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "En arena:"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   17
               Top             =   1680
               Width           =   840
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "VS"
               Height          =   195
               Index           =   1
               Left            =   2400
               TabIndex        =   16
               Top             =   720
               Width           =   255
            End
         End
         Begin VB.CommandButton cmdAgregarParticipantePelea 
            Caption         =   "Agregar participante a pelea"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   2880
            Width           =   5055
         End
         Begin VB.ListBox lstParticipantes 
            Height          =   2205
            ItemData        =   "frmTournament.frx":0038
            Left            =   120
            List            =   "frmTournament.frx":003A
            TabIndex        =   3
            Top             =   480
            Width           =   5055
         End
         Begin VB.Label lblParticipantes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Participantes"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1125
         End
      End
   End
End
Attribute VB_Name = "frmTournament"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum eMapType
    ieInicial
    ieFinal
    ieArena1
    ieArena2
    ieArena3
    ieArena4
    ieArena5
    
    ieLastOption
End Enum

Private bLoading As Boolean
Private iArenaIndex As Integer

Private vbEditChanges() As Boolean
Private vbEditMapChanges() As Boolean
Private vbEditArenaChanges() As Boolean

Private Sub chkEjecutarAlMorir_Click()
    vbEditChanges(eTournamentEdit.ieKillAfterLoose) = True
On Error GoTo ErrHandler
  
    cmdGuardarConfig.Enabled = True
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub chkEjecutarAlMorir_Click de frmTournament.frm")
End Sub

Private Sub cmdGuardarConfig_Click()
'TODO_TORNEO: revisar q se modifico y mandarlo al server
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CloseWindow
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
  
    iArenaIndex = 1
    Call DiscardChanges
    
    ' Combos
    Load_cboMaps
    
    ' Lists
    'TODO_TORNEO: cargar clases e items
    
    Call modCustomCursors.SetFormCursorDefault(Me)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmTournament.frm")
End Sub

Private Sub Load_cboMaps()
    
On Error GoTo ErrHandler
  
    With cboMaps
        .Clear
        .AddItem "Inicial": .ItemData(.NewIndex) = eMapType.ieInicial
        .AddItem "Final": .ItemData(.NewIndex) = eMapType.ieFinal
        .AddItem "Arena 1": .ItemData(.NewIndex) = eMapType.ieArena1
        .AddItem "Arena 2": .ItemData(.NewIndex) = eMapType.ieArena2
        .AddItem "Arena 3": .ItemData(.NewIndex) = eMapType.ieArena3
        .AddItem "Arena 4": .ItemData(.NewIndex) = eMapType.ieArena4
        .AddItem "Arena 5": .ItemData(.NewIndex) = eMapType.ieArena5
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Load_cboMaps de frmTournament.frm")
End Sub

Public Sub Load_lstCompetitors()
On Error GoTo ErrHandler

    lstParticipantes.Clear
    
    With Tournament
        Dim Index As Long
        For Index = 0 To UBound(.CompetitorsList)
            lstParticipantes.AddItem .CompetitorsList(Index)
        Next Index
    End With

ErrHandler:
End Sub

Public Sub UpdateConfig()
On Error GoTo ErrHandler
  

    Dim lTemp As Long
    With Tournament
        ' General
        txtMinNivel.text = .MinLevel
        txtMAxNivel.text = .MaxLevel
        txtNroParticipantes.text = .MaxCompetitors
        txtNumRounds.text = .NumRoundsToWin
        txtOroRequerido.text = .RequiredGold
        chkEjecutarAlMorir.value = .KillAfterLoose
        
        ' Classes
        lstClasesPermitidas.Visible = False
        For lTemp = 0 To lstClasesPermitidas.ListCount - 1
            lstClasesPermitidas.Selected(lTemp) = IsPermitedClass(lstClasesPermitidas.ItemData(lTemp))
        Next lTemp
        lstClasesPermitidas.Visible = True
        
        ' Items
        lstItemsProhibidos.Visible = False
        For lTemp = 0 To lstClasesPermitidas.ListCount - 1
            lstItemsProhibidos.Selected(lTemp) = IsForbiddenITem(lstItemsProhibidos.ItemData(lTemp))
        Next lTemp
        lstItemsProhibidos.Visible = True
    
        ' Maps
        UpdateMapsInfo
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateConfig de frmTournament.frm")
End Sub

Private Function IsPermitedClass(ByVal Clase As Long) As Boolean
On Error GoTo ErrHandler
  
    
    Dim lTemp As Long
    
    With Tournament
        For lTemp = 1 To .NumPermitedClass
            If .PermitedClass(lTemp) = Clase Then
                IsPermitedClass = True
                Exit Function
            End If
        Next lTemp
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function IsPermitedClass de frmTournament.frm")
End Function

Private Function IsForbiddenITem(ByVal ItemIndex As Integer) As Boolean
On Error GoTo ErrHandler
  
    
    Dim lTemp As Long
    
    With Tournament
        For lTemp = 1 To .NumPermitedClass
            If .ForbiddenItem(lTemp) = ItemIndex Then
                IsForbiddenITem = True
                Exit Function
            End If
        Next lTemp
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function IsForbiddenITem de frmTournament.frm")
End Function

Private Sub DiscardChanges()
On Error GoTo ErrHandler
  
    ReDim vbEditChanges(0 To eTournamentEdit.ieLastOption - 1)
    ReDim vbEditMapChanges(0 To eMapType.ieLastOption - 1)
    ReDim vbEditArenaChanges(1 To MAX_ARENAS)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DiscardChanges de frmTournament.frm")
End Sub

Private Sub cmdAgregarPArticipantePelea_Click()
    
On Error GoTo ErrHandler
  
    With lstParticipantes
        If .ListIndex = -1 Then Exit Sub
            
        Dim sParticipante As String
        sParticipante = .List(.ListIndex)
            
        If txtParticipante1.text = sParticipante Or txtParticipante2.text = sParticipante Then
            Call MsgBox("Ya se ingresó el participante!", vbInformation)
            Exit Sub
        End If
        
        If LenB(txtParticipante1.text) = 0 Then
            txtParticipante1.text = sParticipante
            
        ElseIf LenB(txtParticipante2.text) = 0 Then
            txtParticipante1.text = sParticipante
        Else
            ' No llega nunca aca porque se bloquea el boton, pero por las dudas..
            Call MsgBox("Ya se ingresaron los dos participantes!", vbInformation)
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub cmdAgregarPArticipantePelea_Click de frmTournament.frm")
End Sub

Private Sub cmdQuitarParticipante1_Click()
    txtParticipante1.text = vbNullString
End Sub

Private Sub cmdQuitarParticipante2_Click()
    txtParticipante2.text = vbNullString
End Sub

Private Sub lstClasesPermitidas_Click()
On Error GoTo ErrHandler
  
    vbEditChanges(eTournamentEdit.iePermitedClass) = True
    cmdGuardarConfig.Enabled = True
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub lstClasesPermitidas_Click de frmTournament.frm")
End Sub

Private Sub lstItemsProhibidos_Click()

On Error GoTo ErrHandler
  
    With lstItemsProhibidos
        If .ListIndex = -1 Then Exit Sub
        
        If .Selected(.ListIndex) Then
            ' Add
            lstItemsProhibidosSel.AddItem .List(.ListIndex)
            lstItemsProhibidosSel.ItemData(lstItemsProhibidosSel.NewIndex) = .ItemData(.ListIndex)
        Else
            ' Remove
            RemoveForbiddenItem .ItemData(.ListIndex)
        End If
        
        vbEditChanges(eTournamentEdit.ieForbiddenItems) = True
        cmdGuardarConfig.Enabled = True
    
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub lstItemsProhibidos_Click de frmTournament.frm")
End Sub

Private Sub RemoveForbiddenItem(ByVal ObjIndex As Integer)
On Error GoTo ErrHandler
  
    
    Dim lCounter As Long
    With lstItemsProhibidosSel
        For lCounter = 0 To .ListCount - 1
            If .ItemData(lCounter) = ObjIndex Then
                .RemoveItem lCounter
                Exit Sub
            End If
        Next lCounter
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RemoveForbiddenItem de frmTournament.frm")
End Sub

Private Sub optArena_Click(Index As Integer)
    iArenaIndex = Index + 1
End Sub

Private Sub txtMAxNivel_Change()
On Error GoTo ErrHandler
  
    vbEditChanges(eTournamentEdit.ieMaxLevel) = True
    cmdGuardarConfig.Enabled = True
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub txtMAxNivel_Change de frmTournament.frm")
End Sub

Private Sub txtMinNivel_Change()
On Error GoTo ErrHandler
  
    vbEditChanges(eTournamentEdit.ieMinLevel) = True
    cmdGuardarConfig.Enabled = True
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub txtMinNivel_Change de frmTournament.frm")
End Sub

Private Sub txtNroParticipantes_Change()
    vbEditChanges(eTournamentEdit.ieMaxCompetitor) = True
On Error GoTo ErrHandler
  
    cmdGuardarConfig.Enabled = True
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub txtNroParticipantes_Change de frmTournament.frm")
End Sub

Private Sub txtNumRounds_Change()
    vbEditChanges(eTournamentEdit.ieNumRoundsToWin) = True
On Error GoTo ErrHandler
  
    cmdGuardarConfig.Enabled = True
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub txtNumRounds_Change de frmTournament.frm")
End Sub

Private Sub txtOroRequerido_Change()
    vbEditChanges(eTournamentEdit.ieRequiredGold) = True
On Error GoTo ErrHandler
  
    cmdGuardarConfig.Enabled = True
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub txtOroRequerido_Change de frmTournament.frm")
End Sub

Private Sub txtParticipante1_Change()
    EnableFightButtons
End Sub

Private Sub txtParticipante2_Change()
    EnableFightButtons
End Sub

Private Sub EnableFightButtons()
On Error GoTo ErrHandler
  
    Dim bReady As Boolean
    bReady = (LenB(txtParticipante1.text) <> 0 And LenB(txtParticipante2.text) <> 0)
    cmdComenzarPelea.Enabled = bReady
    cmdAgregarParticipantePelea.Enabled = Not bReady
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EnableFightButtons de frmTournament.frm")
End Sub

Private Sub cmdActualizarListaParticipantes_Click()
    Call WriteRequestTournamentCompetitors
End Sub

Private Sub cmdComenzarPelea_Click()
On Error GoTo ErrHandler
  
    Call WriteTournamentFight(txtParticipante1.text, txtParticipante2.text, iArenaIndex)
    Call ResetFightControls
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub cmdComenzarPelea_Click de frmTournament.frm")
End Sub

Private Sub ResetFightControls()
On Error GoTo ErrHandler
  
    txtParticipante1.text = vbNullString
    txtParticipante2.text = vbNullString
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetFightControls de frmTournament.frm")
End Sub

Private Sub cmdConfiguracionActual_Click()
    Call WriteRequestTournamentConfig
End Sub

Private Sub cboMaps_Click()
    UpdateMapsInfo
End Sub

Private Sub UpdateMapsInfo()
On Error GoTo ErrHandler
  
    With cboMaps
        If .ListIndex = -1 Then Exit Sub
        bLoading = True
        LoadMapInfo .ItemData(.ListIndex)
        bLoading = False
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateMapsInfo de frmTournament.frm")
End Sub

Private Sub LoadMapInfo(ByVal iType As Integer)
On Error GoTo ErrHandler
  
    
    Select Case iType
        Case eMapType.ieInicial, eMapType.ieFinal
            picPositions(0).Visible = True
            picPositions(1).Visible = False
            
            If iType = eMapType.ieInicial Then
                With Tournament.WaitingMap
                    txtMap.text = .Map
                    txtPosX.text = .X
                    txtPosY.text = .Y
                End With
            Else
                With Tournament.FinalMap
                    txtMap.text = .Map
                    txtPosX.text = .X
                    txtPosY.text = .Y
                End With
            End If
        
        Case eMapType.ieArena1, eMapType.ieArena2, eMapType.ieArena3, eMapType.ieArena4, eMapType.ieArena5
            picPositions(1).Visible = True
            picPositions(0).Visible = False
            
            Dim ArenaIndex As Byte
            ArenaIndex = GetArenaIndex(iType)
            
            With Tournament.Arenas(ArenaIndex)
                txtMap.text = .Map
                txtUser1X.text = .UserPos1.X
                txtUser1Y.text = .UserPos1.Y
                txtUser2X.text = .UserPos2.X
                txtUser2Y.text = .UserPos2.Y
            End With
    End Select
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadMapInfo de frmTournament.frm")
End Sub

Private Function GetArenaIndex(ByVal iType As Integer) As Byte
On Error GoTo ErrHandler
  
    
    Dim Index As Byte
    Select Case iType
        Case eMapType.ieArena1
            Index = 1
        Case eMapType.ieArena2
            Index = 2
        Case eMapType.ieArena3
            Index = 3
        Case eMapType.ieArena4
            Index = 4
        Case eMapType.ieArena5
            Index = 5
    End Select
    
    GetArenaIndex = Index
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetArenaIndex de frmTournament.frm")
End Function

Private Sub CloseWindow()
On Error GoTo ErrHandler
  
    Unload Me
    If frmMain.Visible Then frmMain.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmTournament.frm")
End Sub

