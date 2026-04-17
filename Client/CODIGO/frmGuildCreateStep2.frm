VERSION 5.00
Begin VB.Form frmGuildCreateStep2 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
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
   Begin VB.TextBox TxtNombreClan 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   2080
      MaxLength       =   25
      TabIndex        =   0
      Text            =   "Nombre de Clan"
      Top             =   3370
      Width           =   1750
   End
   Begin VB.Image ImgCrear 
      Height          =   570
      Left            =   3260
      Top             =   4350
      Width           =   1290
   End
   Begin VB.Image ImgVolver 
      Height          =   570
      Left            =   1430
      Top             =   4350
      Width           =   1290
   End
End
Attribute VB_Name = "frmGuildCreateStep2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Private cButtonReturn As clsGraphicalButton
Private cButtonNext As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()

    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me, , False

    Call LoadControls
    
    TxtNombreClan.text = GuildCreation.Name
    
    Call modCustomCursors.SetFormCursorDefault(Me)
End Sub

Private Sub LoadControls()
    
    Set cButtonReturn = New clsGraphicalButton
    Set cButtonNext = New clsGraphicalButton

    
    Set LastButtonPressed = New clsGraphicalButton
    
    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI
    
    Me.Picture = LoadPicture(GrhPath & "VentanaGuildCreationStep2.jpg")
    
    Call cButtonReturn.Initialize(ImgVolver, GrhPath & "BotonGuildCreationVolver.jpg", _
                                    GrhPath & "BotonGuildCreationVolver.jpg", _
                                    GrhPath & "BotonGuildCreationVolver.jpg", Me)
                                    'GrhPath & "BotonGuildCreationVolver.jpg", _
                                    'GrhPath & "BotonGuildCreationVolver.jpg", Me)
                                    
    Call cButtonNext.Initialize(ImgCrear, GrhPath & "BotonGuildCreationCrear.jpg", _
                                    GrhPath & "BotonGuildCreationCrear.jpg", _
                                    GrhPath & "BotonGuildCreationCrear.jpg", Me)
                                    'GrhPath & "BotonGuildCreationCrear.jpg", _
                                    'GrhPath & "BotonGuildCreationCrear.jpg", Me)

    
End Sub

Private Sub ImgCrear_Click()
    Dim Result As Integer

    If TxtNombreClan.text = "" Then
        Call frmMessageBox.ShowMessage("¡Ingrese un nombre para el clan!")
        Exit Sub
    End If
    
    GuildCreation.Name = Trim(TxtNombreClan.text)
    
    If Len(Trim(TxtNombreClan.Text)) > MAX_GUILD_NAME_LEN Then
        Call frmMessageBox.ShowMessage("El nombre del clan no puede superar los " & MAX_GUILD_NAME_LEN & " caracteres.")
        Exit Sub
    End If
    
    Result = MsgBox("Vas a crear el clan " & GuildCreation.Name & " de alineacion " & GetNameOfAlignment(GuildCreation.Alignment) & ".", vbYesNo, "Creacion de Clan")
        
    If Result = vbYes Then
        Call WriteGuildCreate(TxtNombreClan.text)
    Else
        Me.SetFocus
        Exit Sub
    End If
    
    Unload frmGuildCreateStep1
    Unload frmGuildCreateStep2
    
End Sub

Private Sub ImgVolver_Click()
    frmGuildCreateStep1.Show , frmMain
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CloseWindow
End Sub

Private Sub CloseWindow()
On Error GoTo ErrHandler
  
    Unload Me
    If frmMain.Visible Then frmMain.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmGuildCreateStep2.frm")
End Sub

