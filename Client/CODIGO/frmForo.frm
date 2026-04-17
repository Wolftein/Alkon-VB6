VERSION 5.00
Begin VB.Form frmForo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   457
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   414
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   315
      Left            =   1140
      MaxLength       =   35
      TabIndex        =   2
      Top             =   900
      Visible         =   0   'False
      Width           =   4620
   End
   Begin VB.TextBox txtPost 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   3960
      Left            =   780
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmForo.frx":0000
      Top             =   1935
      Visible         =   0   'False
      Width           =   4770
   End
   Begin VB.ListBox lstTitulos 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   5100
      Left            =   765
      TabIndex        =   0
      Top             =   825
      Width           =   4785
   End
   Begin VB.Label lblTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   1125
      TabIndex        =   4
      Top             =   960
      Width           =   4695
   End
   Begin VB.Image imgMarcoTexto 
      Height          =   465
      Left            =   1095
      Top             =   840
      Width           =   4725
   End
   Begin VB.Image imgCerrar 
      Height          =   255
      Left            =   4080
      Top             =   6060
      Width           =   1455
   End
   Begin VB.Image imgListaMsg 
      Height          =   255
      Left            =   2400
      Top             =   6060
      Width           =   1455
   End
   Begin VB.Image imgDejarMsg 
      Height          =   255
      Left            =   720
      Top             =   6060
      Width           =   1455
   End
   Begin VB.Label lblAutor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
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
      Left            =   1125
      TabIndex        =   3
      Top             =   1455
      Width           =   4650
   End
   Begin VB.Image imgTab 
      Height          =   255
      Index           =   2
      Left            =   4320
      Top             =   360
      Width           =   1575
   End
   Begin VB.Image imgTab 
      Height          =   255
      Index           =   1
      Left            =   2520
      Top             =   360
      Width           =   1575
   End
   Begin VB.Image imgTab 
      Height          =   255
      Index           =   0
      Left            =   960
      Top             =   360
      Width           =   1575
   End
   Begin VB.Image imgDejarAnuncio 
      Height          =   255
      Left            =   2400
      Top             =   6060
      Width           =   1455
   End
End
Attribute VB_Name = "frmForo"
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

Private cBotonDejarAnuncio As clsGraphicalButton
Private cBotonDejarMsg As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton
Private cBotonListaMsg As clsGraphicalButton
Public LastButtonPressed As clsGraphicalButton

' Para controlar las imagenes de fondo y el envio de posteos
Private ForoActual As eForumType
Private VerListaMsg As Boolean
Private Lectura As Boolean

Public ForoLimpio As Boolean
Private Sticky As Boolean

' Para restringir la visibilidad de los foros
Public Privilegios As Byte
Public ForosVisibles As eForumType
Public CanPostSticky As Byte

' Imagenes de fondo
Private FondosDejarMsg(0 To 2) As Picture
Private FondosListaMsg(0 To 2) As Picture

Private Sub Form_Unload(Cancel As Integer)
    MirandoForo = False
On Error GoTo ErrHandler
  
    Privilegios = 0
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Unload de frmForo.frm")
End Sub

Private Sub imgDejarAnuncio_Click()
    Lectura = False
On Error GoTo ErrHandler
  
    VerListaMsg = False
    Sticky = True
    
    'Switch to proper background
    ToogleScreen
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgDejarAnuncio_Click de frmForo.frm")
End Sub

Private Sub imgDejarMsg_Click()
    If Not cBotonDejarMsg.IsEnabled Then Exit Sub
On Error GoTo ErrHandler
  
    
    Dim PostStyle As Byte
    
    If Not VerListaMsg Then
        If Not Lectura Then
        
            If Sticky Then
                PostStyle = GetStickyPost
            Else
                PostStyle = GetNormalPost
            End If

            Call WriteForumPost(txtTitulo.text, txtPost.text, PostStyle)
            
            ' Actualizo localmente
            Call clsForos.AddPost(ForoActual, txtTitulo.text, UserName, txtPost.text, Sticky)
            Call UpdateList
            
            VerListaMsg = True
        End If
    Else
        VerListaMsg = False
        Sticky = False
    End If
    
    Lectura = False
    
    'Switch to proper background
    ToogleScreen
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgDejarMsg_Click de frmForo.frm")
End Sub

Private Sub imgCerrar_Click()
    CerrarVentana
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CerrarVentana
End Sub

Private Sub CerrarVentana()
On Error GoTo ErrHandler
  
    Unload Me
    If frmMain.Visible Then frmMain.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CerrarVentana de frmForo.frm")
End Sub

Private Sub imgListaMsg_Click()
On Error GoTo ErrHandler
  
    VerListaMsg = True
    ToogleScreen
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgListaMsg_Click de frmForo.frm")
End Sub


Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
On Error GoTo ErrHandler
  
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Call LoadControls
    
    ' Initial config
    ForoActual = eForumType.ieGeneral
    VerListaMsg = True
    UpdateList
    
    ' Default background
    ToogleScreen
    
    ForoLimpio = False
    MirandoForo = True
    
    ' Si no es caos o gms, no puede ver el tab de caos.
    If (Privilegios And eForumVisibility.ieCAOS_MEMBER) = 0 Then imgTab(2).Visible = False
    
    ' Si no es armada o gm, no puede ver el tab de armadas.
    If (Privilegios And eForumVisibility.ieREAL_MEMBER) = 0 Then imgTab(1).Visible = False
    
    Call modCustomCursors.SetFormCursorDefault(Me)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmForo.frm")
End Sub

Private Sub LoadControls()
On Error GoTo ErrHandler
  
    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI
    
    ' Load pictures
    Set FondosListaMsg(eForumType.ieGeneral) = LoadPicture(GrhPath & "ForoGeneral.jpg")
    Set FondosListaMsg(eForumType.ieREAL) = LoadPicture(GrhPath & "ForoReal.jpg")
    Set FondosListaMsg(eForumType.ieCAOS) = LoadPicture(GrhPath & "ForoCaos.jpg")
    
    Set FondosDejarMsg(eForumType.ieGeneral) = LoadPicture(GrhPath & "ForoMsgGeneral.jpg")
    Set FondosDejarMsg(eForumType.ieREAL) = LoadPicture(GrhPath & "ForoMsgReal.jpg")
    Set FondosDejarMsg(eForumType.ieCAOS) = LoadPicture(GrhPath & "ForoMsgCaos.jpg")
    
    imgMarcoTexto.Picture = LoadPicture(GrhPath & "MarcoTextBox.jpg")

    Set cBotonDejarAnuncio = New clsGraphicalButton
    Set cBotonDejarMsg = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonListaMsg = New clsGraphicalButton

    Set LastButtonPressed = New clsGraphicalButton

    ' Initialize buttons
    Call cBotonDejarAnuncio.Initialize(imgDejarAnuncio, GrhPath & "BotonDejarAnuncioForo.jpg", _
                                            GrhPath & "BotonDejarAnuncioRolloverForo.jpg", _
                                            GrhPath & "BotonDejarAnuncioClickForo.jpg", Me)
                                            
    Call cBotonDejarMsg.Initialize(imgDejarMsg, GrhPath & "BotonDejarMsgForo.jpg", _
                                            GrhPath & "BotonDejarMsgRolloverForo.jpg", _
                                            GrhPath & "BotonDejarMsgClickForo.jpg", Me, _
                                            GrhPath & "BotonDejarMsgDisabledForo.jpg")
                                            
    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCerrarForo.jpg", _
                                            GrhPath & "BotonCerrarRolloverForo.jpg", _
                                            GrhPath & "BotonCerrarClickForo.jpg", Me)
                                            
    Call cBotonListaMsg.Initialize(imgListaMsg, GrhPath & "BotonListaMsgForo.jpg", _
                                            GrhPath & "BotonListaMsgRolloverForo.jpg", _
                                            GrhPath & "BotonListaMsgClickForo.jpg", Me)

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadControls de frmForo.frm")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgTab_Click(Index As Integer)

On Error GoTo ErrHandler
  
    Call Engine_Audio.PlayInterface(SND_CLICK)
    
    If Index <> ForoActual Then
        ForoActual = Index
        VerListaMsg = True
        Lectura = False
        UpdateList
        ToogleScreen
    Else
        If Not VerListaMsg Then
            VerListaMsg = True
            Lectura = False
            ToogleScreen
        End If
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgTab_Click de frmForo.frm")
End Sub

Private Sub ToogleScreen()
On Error GoTo ErrHandler
  
    
    Dim PostOffset As Integer
    
    imgMarcoTexto.Visible = Not VerListaMsg And Not Lectura
    txtTitulo.Visible = Not VerListaMsg And Not Lectura
    lblTitulo.Visible = Not VerListaMsg And Lectura
    
    Call cBotonDejarMsg.EnableButton(VerListaMsg Or Lectura)
    
    txtPost.Visible = Not VerListaMsg
    
    imgDejarAnuncio.Visible = VerListaMsg And PuedeDejarAnuncios
    imgListaMsg.Visible = Not VerListaMsg
    lstTitulos.Visible = VerListaMsg
    
    If VerListaMsg Then
        Me.Picture = FondosListaMsg(ForoActual)
    Else
        If Lectura Then
            With lstTitulos
                PostOffset = .ItemData(.ListIndex)
                
                ' Normal post?
                If PostOffset < STICKY_FORUM_OFFSET Then
                    lblTitulo.Caption = Foros(ForoActual).GeneralTitle(PostOffset)
                    txtPost.text = Foros(ForoActual).GeneralPost(PostOffset)
                    lblAutor.Caption = Foros(ForoActual).GeneralAuthor(PostOffset)
                
                ' Sticky post
                Else
                    PostOffset = PostOffset - STICKY_FORUM_OFFSET
                    
                    lblTitulo.Caption = Foros(ForoActual).StickyTitle(PostOffset)
                    txtPost.text = Foros(ForoActual).StickyPost(PostOffset)
                    lblAutor.Caption = Foros(ForoActual).StickyAuthor(PostOffset)
                End If
            End With
        Else
            lblAutor.Caption = UserName
            txtTitulo.text = vbNullString
            txtPost.text = vbNullString
            
            txtTitulo.SetFocus
        End If
        
        txtPost.Locked = Lectura
        Me.Picture = FondosDejarMsg(ForoActual)
    End If
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ToogleScreen de frmForo.frm")
End Sub

Private Function PuedeDejarAnuncios() As Boolean
On Error GoTo ErrHandler
  
    
    ' No puede
    If CanPostSticky = 0 Then Exit Function

    If ForoActual = eForumType.ieGeneral Then
        ' Solo puede dejar en el general si es gm
        If CanPostSticky <> 2 Then Exit Function
    End If
    
    PuedeDejarAnuncios = True
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PuedeDejarAnuncios de frmForo.frm")
End Function

Private Sub lstTitulos_Click()
On Error GoTo ErrHandler
  
    VerListaMsg = False
    Lectura = True
    ToogleScreen
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub lstTitulos_Click de frmForo.frm")
End Sub

Private Sub lstTitulos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub txtPost_Change()
On Error GoTo ErrHandler
  
    If Lectura Then Exit Sub
    
    Call cBotonDejarMsg.EnableButton(Len(txtTitulo.text) <> 0 And Len(txtPost.text) <> 0)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub txtPost_Change de frmForo.frm")
End Sub

Private Sub txtPost_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub txtTitulo_Change()
On Error GoTo ErrHandler
  
    If Lectura Then Exit Sub
    
    Call cBotonDejarMsg.EnableButton(Len(txtTitulo.text) <> 0 And Len(txtPost.text) <> 0)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub txtTitulo_Change de frmForo.frm")
End Sub

Private Sub UpdateList()
On Error GoTo ErrHandler
  
    Dim PostIndex As Long
    
    lstTitulos.Clear
    
    With lstTitulos
        ' Sticky first
        For PostIndex = 1 To clsForos.GetNroSticky(ForoActual)
            .AddItem "[ANUNCIO] " & Foros(ForoActual).StickyTitle(PostIndex) & " (" & Foros(ForoActual).StickyAuthor(PostIndex) & ")"
            .ItemData(.NewIndex) = STICKY_FORUM_OFFSET + PostIndex
        Next PostIndex
    
        ' Then normal posts
        For PostIndex = 1 To clsForos.GetNroPost(ForoActual)
            .AddItem Foros(ForoActual).GeneralTitle(PostIndex) & " (" & Foros(ForoActual).GeneralAuthor(PostIndex) & ")"
            .ItemData(.NewIndex) = PostIndex
        Next PostIndex
    End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateList de frmForo.frm")
End Sub

Private Function GetStickyPost() As Byte
On Error GoTo ErrHandler
  
    Select Case ForoActual
        Case 0
            GetStickyPost = eForumMsgType.ieGENERAL_STICKY
            
        Case 1
            GetStickyPost = eForumMsgType.ieREAL_STICKY
            
        Case 2
            GetStickyPost = eForumMsgType.ieCAOS_STICKY
            
    End Select
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetStickyPost de frmForo.frm")
End Function

Private Function GetNormalPost() As Byte
On Error GoTo ErrHandler
  
    Select Case ForoActual
        Case 0
            GetNormalPost = eForumMsgType.ieGeneral
            
        Case 1
            GetNormalPost = eForumMsgType.ieREAL
            
        Case 2
            GetNormalPost = eForumMsgType.ieCAOS
            
    End Select
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetNormalPost de frmForo.frm")
End Function
