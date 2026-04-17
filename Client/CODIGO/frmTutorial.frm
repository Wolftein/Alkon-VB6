VERSION 5.00
Begin VB.Form frmTutorial 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   583
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgSalir 
      Height          =   195
      Left            =   8400
      Top             =   150
      Width           =   195
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Esto es un título de prueba"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   390
      Left            =   525
      TabIndex        =   3
      Top             =   645
      Width           =   7725
   End
   Begin VB.Image imgCheck 
      Height          =   450
      Left            =   3060
      Top             =   6900
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgMostrar 
      Height          =   570
      Left            =   3000
      Top             =   6855
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Image imgSiguiente 
      Height          =   375
      Left            =   5055
      Top             =   6990
      Width           =   2775
   End
   Begin VB.Image imgAnterior 
      Height          =   375
      Left            =   915
      Top             =   6990
      Width           =   2775
   End
   Begin VB.Label lblMensaje 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTutorial.frx":0000
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
      Height          =   5550
      Left            =   585
      TabIndex        =   2
      Top             =   1125
      Width           =   7590
   End
   Begin VB.Label lblPagTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10"
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
      Left            =   7980
      TabIndex        =   1
      Top             =   225
      Width           =   255
   End
   Begin VB.Label lblPagActual 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
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
      Left            =   7470
      TabIndex        =   0
      Top             =   225
      Width           =   255
   End
End
Attribute VB_Name = "frmTutorial"
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
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonSiguiente As clsGraphicalButton
Private cBotonAnterior As clsGraphicalButton
Private cBotonSalir As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Type tTutorial
    sTitle As String
    sPage As String
End Type

Private picCheck As Picture
Private picMostrar As Picture

Private Tutorial() As tTutorial
Private NumPages As Long
Private CurrentPage As Long

Private Sub Form_Load()

    ' Handles Form movement (drag and drop).
On Error GoTo ErrHandler
  
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "VentanaTutorial.jpg")
    
    Call LoadButtons
    
    Call LoadTutorial
    
    CurrentPage = 1
    Call SelectPage(CurrentPage)
    
    Call modCustomCursors.SetFormCursorDefault(Me)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmTutorial.frm")
End Sub

Private Sub LoadButtons()
On Error GoTo ErrHandler
  
    Dim GrhPath As String
    
    GrhPath = DirInterfaces & SELECTED_UI

    Set cBotonSiguiente = New clsGraphicalButton
    Set cBotonAnterior = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    
    
    Set LastButtonPressed = New clsGraphicalButton
    
    
    Call cBotonSiguiente.Initialize(imgSiguiente, GrhPath & "BotonTutorialSiguiente.jpg", _
                                    GrhPath & "BotonTutorialSiguienteRollover.jpg", _
                                    GrhPath & "BotonTutorialSiguienteClick.jpg", Me, _
                                    GrhPath & "BotonTutorialSiguienteDisabled.jpg")

    Call cBotonAnterior.Initialize(imgAnterior, GrhPath & "BotonTutorialAnterior.jpg", _
                                    GrhPath & "BotonTutorialAnteriorRollover.jpg", _
                                    GrhPath & "BotonTutorialAnteriorClick.jpg", Me, _
                                    GrhPath & "BotonTutorialAnteriorDisabled.jpg", True)
                                    
    Call cBotonSalir.Initialize(imgSalir, GrhPath & "BotonCerrarForm.jpg", _
                                    GrhPath & "BotonCerrarFormRollover.jpg", _
                                    GrhPath & "BotonCerrarFormClick.jpg", Me)
                                    
    'Set picCheck = LoadPicture(GrhPath & "CheckTutorial.jpg")
    'Set picMostrar = LoadPicture(GrhPath & "NoMostrarTutorial.jpg")
    
    'imgMostrar.Picture = picMostrar
    
    'If Not bShowTutorial Then
    '    imgCheck.Picture = picCheck
    'Else
    '    Set imgCheck.Picture = Nothing
    'End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadButtons de frmTutorial.frm")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub


Private Sub imgAnterior_Click()

On Error GoTo ErrHandler
  
    If Not cBotonAnterior.IsEnabled Then Exit Sub
    
    CurrentPage = CurrentPage - 1
    
    If CurrentPage = 1 Then Call cBotonAnterior.EnableButton(False)
    
    If Not cBotonSiguiente.IsEnabled Then Call cBotonSiguiente.EnableButton(True)
    
    Call SelectPage(CurrentPage)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgAnterior_Click de frmTutorial.frm")
End Sub

Private Sub imgCheck_Click()
    
On Error GoTo ErrHandler
  
    bShowTutorial = Not bShowTutorial
    
    If Not bShowTutorial Then
        imgCheck.Picture = picCheck
    Else
        Set imgCheck.Picture = Nothing
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgCheck_Click de frmTutorial.frm")
End Sub

Private Sub imgSalir_Click()
    Call CerrarVentana
End Sub

Private Sub ImgSiguiente_Click()
    
On Error GoTo ErrHandler
  
    If Not cBotonSiguiente.IsEnabled Then Exit Sub
    
    CurrentPage = CurrentPage + 1
    
    ' DEshabilita el boton siguiente si esta en la ultima pagina
    If CurrentPage = NumPages Then Call cBotonSiguiente.EnableButton(False)
    
    ' Habilita el boton anterior
    If Not cBotonAnterior.IsEnabled Then Call cBotonAnterior.EnableButton(True)
    
    Call SelectPage(CurrentPage)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgSiguiente_Click de frmTutorial.frm")
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CerrarVentana
End Sub

Private Sub CerrarVentana()
On Error GoTo ErrHandler
  
    bShowTutorial = False 'Mientras no se pueda tildar/destildar para verlo más tarde, esto queda así :P
    Unload Me
    If frmMain.Visible Then frmMain.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CerrarVentana de frmTutorial.frm")
End Sub

Private Sub LoadTutorial()
On Error GoTo ErrHandler
  
    
    Dim TutorialPath As String
    Dim lPage As Long
    Dim NumLines As Long
    Dim lLine As Long
    Dim sLine As String
    
    TutorialPath = App.path & DAT_PATH & "Tutorial.dat"
    NumPages = Val(GetVar(TutorialPath, "INIT", "NumPags"))
    
    If NumPages > 0 Then
        ReDim Tutorial(1 To NumPages)
        
        ' Cargo paginas
        For lPage = 1 To NumPages
            NumLines = Val(GetVar(TutorialPath, "PAG" & lPage, "NumLines"))
            
            With Tutorial(lPage)
                
                .sTitle = GetVar(TutorialPath, "PAG" & lPage, "Title")
                
                ' Cargo cada linea de la pagina
                For lLine = 1 To NumLines
                    sLine = GetVar(TutorialPath, "PAG" & lPage, "Line" & lLine)
                    .sPage = .sPage & sLine & vbCrLf
                Next lLine
            End With
            
        Next lPage
    End If
    
    lblPagTotal.Caption = NumPages
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadTutorial de frmTutorial.frm")
End Sub

Private Sub SelectPage(ByVal lPage As Long)
On Error GoTo ErrHandler
  
    lblTitulo.Caption = Tutorial(lPage).sTitle
    lblMensaje.Caption = Tutorial(lPage).sPage
    lblPagActual.Caption = lPage
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SelectPage de frmTutorial.frm")
End Sub


Private Sub lblMensaje_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub
