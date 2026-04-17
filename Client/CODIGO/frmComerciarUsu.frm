VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Begin VB.Form frmComerciarUsu 
   BorderStyle     =   0  'None
   ClientHeight    =   8850
   ClientLeft      =   7605
   ClientTop       =   2490
   ClientWidth     =   9975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   590
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ARGENTUM.AOPictureBox picInvOfertaOtro 
      Height          =   2880
      Left            =   6960
      TabIndex        =   5
      Top             =   4695
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   5080
   End
   Begin ARGENTUM.AOPictureBox picInvOfertaProp 
      Height          =   2880
      Left            =   6960
      TabIndex        =   4
      Top             =   945
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   5080
   End
   Begin ARGENTUM.AOPictureBox picInvComercio 
      Height          =   2880
      Left            =   630
      TabIndex        =   3
      Top             =   945
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   5080
   End
   Begin VB.TextBox txtAgregar 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   270
      Left            =   4410
      TabIndex        =   1
      Text            =   "1"
      Top             =   2385
      Width           =   1140
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Left            =   480
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   7350
      Width           =   5790
   End
   Begin RichTextLib.RichTextBox CommerceConsole 
      Height          =   1575
      Left            =   480
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   5400
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   2778
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmComerciarUsu.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgSendMessage 
      Height          =   270
      Left            =   6270
      Top             =   7350
      Width           =   270
   End
   Begin VB.Image imgWhitdrawGold 
      Height          =   270
      Left            =   6000
      Top             =   975
      Width           =   270
   End
   Begin VB.Image imgAddGold 
      Height          =   270
      Left            =   3735
      Top             =   975
      Width           =   270
   End
   Begin VB.Label lblOtherUserGold 
      BackStyle       =   0  'Transparent
      Caption         =   "100000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   7800
      TabIndex        =   8
      Top             =   7770
      Width           =   1575
   End
   Begin VB.Label lblUserOfferedGold 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "100000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblUserGold 
      BackStyle       =   0  'Transparent
      Caption         =   "100000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Image imgCancelar 
      Height          =   585
      Left            =   3330
      Tag             =   "1"
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Image imgRechazar 
      Height          =   585
      Left            =   3600
      Tag             =   "2"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgConfirmar 
      Height          =   585
      Left            =   4920
      Tag             =   "2"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgAceptar 
      Height          =   585
      Left            =   5055
      Tag             =   "2"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgAgregar 
      Height          =   270
      Left            =   3375
      Top             =   2400
      Width           =   270
   End
   Begin VB.Image imgQuitar 
      Height          =   270
      Left            =   6375
      Top             =   2370
      Width           =   270
   End
End
Attribute VB_Name = "frmComerciarUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmComerciarUsu.frm
'
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonAceptar As clsGraphicalButton
Private cBotonCancelar As clsGraphicalButton
Private cBotonRechazar As clsGraphicalButton
Private cBotonConfirmar As clsGraphicalButton
Private cButtonAddGold As clsGraphicalButton
Private cButtonWhitdrawGold As clsGraphicalButton
Private cButtonSendMessage As clsGraphicalButton
Private cButtonAddItem As clsGraphicalButton
Private cButtonWhitdrawItem As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private sCommerceChat As String

Private WithEvents dragUsrToConfirm As clsGraphicalInventory
Attribute dragUsrToConfirm.VB_VarHelpID = -1
Private WithEvents dragUsrGoldToConfirm As clsGraphicalInventory
Attribute dragUsrGoldToConfirm.VB_VarHelpID = -1

Private WithEvents dragConfirmToUsrUsr As clsGraphicalInventory
Attribute dragConfirmToUsrUsr.VB_VarHelpID = -1
Private WithEvents dragConfirmToUsrGold As clsGraphicalInventory
Attribute dragConfirmToUsrGold.VB_VarHelpID = -1

Private OfferedGold As Double
Private OtherPlayerOfferedGold As Double

Private Sub CancelAllDrags()
On Error GoTo ErrHandler
  
    'InvComUsu.DragFinish
    'InvOfferComUsu(0).DragFinish
    'InvOfferComUsu(1).DragFinish
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CancelAllDrags de frmComerciarUsu.frm")
End Sub
Private Sub dragUsrToConfirm_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer, _
                                      ByVal moveType As eMoveType)
On Error GoTo ErrHandler
  
    CancelAllDrags

    If moveType = eMoveType.InventoryToTarget Then
    Call InvOfferComUsu(0).DeselectItem
        Call imgAgregar_Click
    End If

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub dragUsrToConfirm_dragDone de frmComerciarUsu.frm")
End Sub
Private Sub dragUsrGoldToConfirm_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer, _
                                      ByVal moveType As eMoveType)
On Error GoTo ErrHandler
  
    CancelAllDrags

    If moveType = eMoveType.InventoryToTarget Then
    InvComUsu.SelectGold
        Call imgAgregar_Click
    End If
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub dragUsrGoldToConfirm_dragDone de frmComerciarUsu.frm")
End Sub
Private Sub dragConfirmToUsrGold_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer, _
                                      ByVal moveType As eMoveType)
On Error GoTo ErrHandler
  
    CancelAllDrags

    If moveType = eMoveType.InventoryToTarget Then
        InvOfferComUsu(0).SelectGold
        Call imgQuitar_Click
    End If
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub dragConfirmToUsrGold_dragDone de frmComerciarUsu.frm")
End Sub


Private Sub dragConfirmToUsrUsr_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer, _
                                      ByVal moveType As eMoveType)
On Error GoTo ErrHandler
  
    CancelAllDrags

    If moveType = eMoveType.InventoryToTarget Then
            InvOfferComUsu(1).DeselectItem
        Call imgQuitar_Click
    End If
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub dragConfirmToUsrUsr_dragDone de frmComerciarUsu.frm")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandler
  
    Set dragConfirmToUsrUsr = Nothing
    Set dragConfirmToUsrGold = Nothing

    Set dragUsrGoldToConfirm = Nothing
    Set dragUsrToConfirm = Nothing
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Unload de frmComerciarUsu.frm")
End Sub

Private Sub imgAceptar_Click()
On Error GoTo ErrHandler
  
    If Not cBotonAceptar.IsEnabled Then Exit Sub  ' Deshabilitado
    
    Call WriteUserCommerceOk
    Call HabilitarAceptarRechazar(False)
    
    Call PrintCommerceMsg("Aceptaste la oferta. Espera a que la otra parte haga lo mismo.", FontTypeNames.FONTTYPE_GUILD)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgAceptar_Click de frmComerciarUsu.frm")
End Sub

Private Sub imgWhitdrawGold_Click()
    On Error GoTo ErrHandler
    Dim Amount As Long
    
    Amount = Val(txtAgregar.Text)
    
    If Amount <= 0 Or OfferedGold <= 0 Then Exit Sub
    
    If Amount > OfferedGold Then Amount = OfferedGold
    
    Amount = Amount
    
    OfferedGold = OfferedGold - Amount

    ' Le aviso al otro de mi cambio de oferta
    Call WriteUserCommerceOfferGold(Amount * -1)
    
    lblUserOfferedGold.Caption = OfferedGold
    LblUserGold.Caption = Val(LblUserGold.Caption) + Amount
    
    If Not HasAnyItem(InvOfferComUsu(0)) And OfferedGold <= 0 Then HabilitarConfirmar (False)
    
    Call PrintCommerceMsg("Retiraste " & Amount & " moneda" & IIf(Amount = 1, "", "s") & " de oro de tu oferta.", FontTypeNames.FONTTYPE_GUILD)


    Exit Sub
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgWhitdrawGold_Click de frmComerciarUsu.frm")
End Sub

Private Sub imgAddGold_Click()
On Error GoTo ErrHandler
    Dim Amount As Long
    
    Amount = Val(txtAgregar.Text)
    
    If (Amount + OfferedGold) > UserGLD Then
        Call PrintCommerceMsg("¡No tienes esa cantidad!", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
    End If
    
    If Amount <= 0 Then Exit Sub
    
    OfferedGold = OfferedGold + Amount

    ' Le aviso al otro de mi cambio de oferta
    Call WriteUserCommerceOfferGold(Amount)
    
    lblUserOfferedGold.Caption = OfferedGold
    LblUserGold.Caption = Val(LblUserGold.Caption) - Amount
    
    Call HabilitarConfirmar(True)
    
    Call PrintCommerceMsg("Agregaste " & Amount & " moneda" & IIf(Amount = 1, "", "s") & " de oro a tu oferta.", FontTypeNames.FONTTYPE_GUILD)


    Exit Sub
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgAddGold_Click de frmComerciarUsu.frm")
End Sub

Private Sub imgSendMessage_Click()
    Call SendChatText(sCommerceChat)
End Sub


Private Sub imgAgregar_Click()
   
    ' No tiene seleccionado ningun item
On Error GoTo ErrHandler
  
    If InvComUsu.SelectedItem = 0 Then
        Call PrintCommerceMsg("¡No tienes ningún item seleccionado!", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
    End If
    
    ' Numero invalido
    If Not IsNumeric(txtAgregar.text) Then Exit Sub
    
    Dim OfferSlot As Byte
    Dim Amount As Long
    Dim InvSlot As Byte
        
    With InvComUsu
        If .SelectedItem > 0 Then
             If Val(txtAgregar.text) > .Amount(.SelectedItem) Then
                Call PrintCommerceMsg("¡No tienes esa cantidad!", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
             
            OfferSlot = CheckAvailableSlot(.SelectedItem, Val(txtAgregar.text))
            
            ' Hay espacio o lugar donde sumarlo?
            If OfferSlot > 0 Then
                If IsSecondaryArmour(.ObjIndex(OfferSlot)) Then
                    Call PrintCommerceMsg("No puedes comerciar este ítem.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
  
                ' Le aviso al otro de mi cambio de oferta
                Call WriteUserCommerceOffer(.SelectedItem, Val(txtAgregar.text), OfferSlot)
                
                ' Actualizo el inventario general de comercio
                Call .ChangeSlotItemAmount(.SelectedItem, .Amount(.SelectedItem) - Val(txtAgregar.text))
                
                Amount = InvOfferComUsu(0).Amount(OfferSlot) + Val(txtAgregar.text)
                
                ' Actualizo los inventarios
                If InvOfferComUsu(0).ObjIndex(OfferSlot) > 0 Then
                    ' Si ya esta el item, solo actualizo su cantidad en el invenatario
                    Call InvOfferComUsu(0).ChangeSlotItemAmount(OfferSlot, Amount)
                Else
                    InvSlot = .SelectedItem
                    ' Si no agrego todo
                    Call InvOfferComUsu(0).SetItem(OfferSlot, .ObjIndex(InvSlot), _
                                                    Amount, 0, .GrhIndex(InvSlot), .OBJType(InvSlot), _
                                                    .MaxHit(InvSlot), .MinHit(InvSlot), .MaxDef(InvSlot), .MinDef(InvSlot), _
                                                    .Valor(InvSlot), .ItemName(InvSlot), 0, .CanUse(InvSlot))
                End If
            End If
        End If
    End With
    
    HabilitarConfirmar True
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgAgregar_Click de frmComerciarUsu.frm")
End Sub

Private Sub imgCancelar_Click()
    Call WriteUserCommerceEnd
End Sub

Private Sub imgConfirmar_Click()
On Error GoTo ErrHandler
  
    If Not cBotonConfirmar.IsEnabled Then Exit Sub  ' Deshabilitado
    
    HabilitarConfirmar False
    imgAgregar.Visible = False
    imgQuitar.Visible = False
    imgAddGold.Visible = False
    txtAgregar.Enabled = False
    
    
    Call PrintCommerceMsg("¡Has confirmado tu oferta! Ya no puedes cambiarla.", FontTypeNames.FONTTYPE_CONSE)
    Call WriteUserCommerceConfirm
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgConfirmar_Click de frmComerciarUsu.frm")
End Sub



Private Sub imgQuitar_Click()
On Error GoTo ErrHandler
  
    Dim Amount As Long

    ' No tiene seleccionado ningun item
    If InvOfferComUsu(0).SelectedItem = 0 Then
        Call PrintCommerceMsg("¡No tienes ningún ítem seleccionado!", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
    End If
    
    ' Numero invalido
    If Not IsNumeric(txtAgregar.text) Then Exit Sub

    Amount = IIf(Val(txtAgregar.Text) > InvOfferComUsu(0).Amount(InvOfferComUsu(0).SelectedItem), _
                InvOfferComUsu(0).Amount(InvOfferComUsu(0).SelectedItem), Val(txtAgregar.Text))
    ' Estoy quitando, paso un valor negativo
    Amount = Amount * (-1)
    
    ' No tiene sentido que se quiten 0 unidades
    If Amount <> 0 Then
        With InvOfferComUsu(0)
            
            Call PrintCommerceMsg("¡¡Quitaste " & Amount * (-1) & " " & .ItemName(.SelectedItem) & " de tu oferta!!", FontTypeNames.FONTTYPE_GUILD)

            ' Le aviso al otro de mi cambio de oferta
            Call WriteUserCommerceOffer(0, Amount, .SelectedItem)
        
            ' Actualizo el inventario general
            Call UpdateInvCom(.ObjIndex(.SelectedItem), Abs(Amount))
             
             ' Actualizo el inventario de oferta
             If .Amount(.SelectedItem) + Amount = 0 Then
                 ' Borro el item
                 Call .SetItem(.SelectedItem, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0, True)
             Else
                 ' Le resto la cantidad deseada
                 Call .ChangeSlotItemAmount(.SelectedItem, .Amount(.SelectedItem) + Amount)
             End If
        End With
    End If

    
    ' Si quito todos los items de la oferta, no puede confirmarla
    If Not HasAnyItem(InvOfferComUsu(0)) And _
       OfferedGold <= 0 Then HabilitarConfirmar (False)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgQuitar_Click de frmComerciarUsu.frm")
End Sub

Private Sub imgRechazar_Click()
On Error GoTo ErrHandler
  
    If Not cBotonRechazar.IsEnabled Then Exit Sub  ' Deshabilitado
    
    Call WriteUserCommerceReject
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgRechazar_Click de frmComerciarUsu.frm")
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
On Error GoTo ErrHandler
  
    OfferedGold = 0
    OtherPlayerOfferedGold = 0
    
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Set InvComUsu.dropInventory = InvOfferComUsu(0)
    Set dragUsrToConfirm = InvComUsu
    
    Set InvOfferComUsu(0).dropInventory = InvComUsu
    Set dragConfirmToUsrUsr = InvOfferComUsu(0)
        
    Call LoadControls
    
    Call modCustomCursors.SetFormCursorDefault(Me)
    
    Call PrintCommerceMsg("> Una vez termines de formar tu oferta, debes presionar en ""Confirmar"", tras lo cual ya no podrás modificarla.", FontTypeNames.FONTTYPE_GUILDMSG)
    Call PrintCommerceMsg("> Luego que el otro usuario confirme su oferta, podrás aceptarla o rechazarla. Si la rechazas, se terminará el comercio.", FontTypeNames.FONTTYPE_GUILDMSG)
    Call PrintCommerceMsg("> Cuando ambos acepten la oferta del otro, se realizará el intercambio.", FontTypeNames.FONTTYPE_GUILDMSG)
    Call PrintCommerceMsg("> Si se intercambian más ítems de los que pueden entrar en tu inventario, es probable que caigan al suelo, así que presta mucha atención a esto.", FontTypeNames.FONTTYPE_GUILDMSG)

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmComerciarUsu.frm")
End Sub

Private Sub LoadControls()
On Error GoTo ErrHandler
  
    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI

    Me.Picture = LoadPicture(GrhPath & "VentanaComercioUsuario.jpg")
    
    
    Set cBotonAceptar = New clsGraphicalButton
    Set cBotonConfirmar = New clsGraphicalButton
    Set cBotonRechazar = New clsGraphicalButton
    Set cBotonCancelar = New clsGraphicalButton
    Set cButtonAddGold = New clsGraphicalButton
    Set cButtonWhitdrawGold = New clsGraphicalButton
    Set cButtonSendMessage = New clsGraphicalButton
    Set cButtonAddItem = New clsGraphicalButton
    Set cButtonWhitdrawItem = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonAceptar.Initialize(imgAceptar, GrhPath & "BotonAceptar.jpg", _
                                        GrhPath & "BotonAceptar.jpg", _
                                        GrhPath & "BotonAceptar.jpg", Me, _
                                        GrhPath & "BotonAceptar.jpg", True)
                                    
    Call cBotonConfirmar.Initialize(imgConfirmar, GrhPath & "BotonConfirmar.jpg", _
                                        GrhPath & "BotonConfirmar.jpg", _
                                        GrhPath & "BotonConfirmar.jpg", Me, _
                                        GrhPath & "BotonConfirmar.jpg", True)
                                        
    Call cBotonRechazar.Initialize(imgRechazar, GrhPath & "BotonRechazar.jpg", _
                                        GrhPath & "BotonRechazar.jpg", _
                                        GrhPath & "BotonRechazar.jpg", Me, _
                                        GrhPath & "BotonRechazar.jpg", True)
                                        
    Call cBotonCancelar.Initialize(ImgCancelar, GrhPath & "BotonCancelar.jpg", _
                                        GrhPath & "BotonCancelar.jpg", _
                                        GrhPath & "BotonCancelar.jpg", Me)
                                        
    Call cButtonAddGold.Initialize(imgAddGold, GrhPath & "BotonFlechaDerecha_2.jpg", _
                                        GrhPath & "BotonFlechaDerecha_2.jpg", _
                                        GrhPath & "BotonFlechaDerecha_2.jpg", Me)
                                        
                                        
    Call cButtonWhitdrawGold.Initialize(imgWhitdrawGold, GrhPath & "BotonFlechaIzquierda_2.jpg", _
                                        GrhPath & "BotonFlechaIzquierda_2.jpg", _
                                        GrhPath & "BotonFlechaIzquierda_2.jpg", Me)
                                        
    Call cButtonSendMessage.Initialize(imgSendMessage, GrhPath & "BotonFlechaDerecha_2.jpg", _
                                        GrhPath & "BotonFlechaDerecha_2.jpg", _
                                        GrhPath & "BotonFlechaDerecha_2.jpg", Me)
                                        
    Call cButtonAddItem.Initialize(imgAgregar, GrhPath & "BotonFlechaDerecha_2.jpg", _
                                        GrhPath & "BotonFlechaDerecha_2.jpg", _
                                        GrhPath & "BotonFlechaDerecha_2.jpg", Me)
                                        
    Call cButtonWhitdrawItem.Initialize(imgQuitar, GrhPath & "BotonFlechaIzquierda_2.jpg", _
                                        GrhPath & "BotonFlechaIzquierda_2.jpg", _
                                        GrhPath & "BotonFlechaIzquierda_2.jpg", Me)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadButtons de frmComerciarUsu.frm")
End Sub

Private Sub Form_LostFocus()
    Me.SetFocus
End Sub

Public Sub UserConfirmedOffer()
    Call HabilitarAceptarRechazar(True)
    Call PrintCommerceMsg(TradingUserName & " ha confirmado su oferta!", FontTypeNames.FONTTYPE_CONSE)
End Sub

Private Sub picInvComercio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub picInvOfertaOtro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub picInvOfertaProp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 03/10/2009
'**************************************************************
On Error GoTo ErrHandler
  
    If Len(SendTxt.text) > 160 Then
        sCommerceChat = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim I As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For I = 1 To Len(SendTxt.text)
            CharAscii = Asc(mid$(SendTxt.text, I, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next I
        
        If tempstr <> SendTxt.text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.text = tempstr
        End If
        
        sCommerceChat = SendTxt.text
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendTxt_Change de frmComerciarUsu.frm")
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
On Error GoTo ErrHandler
  
    If KeyCode = vbKeyReturn Then
        Call SendChatText(sCommerceChat)
        
        KeyCode = 0
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendTxt_KeyUp de frmComerciarUsu.frm")
End Sub


Public Sub SendChatText(ByRef Text As String)
    If LenB(Text) <> 0 Then Call WriteCommerceChat(Text)
    
    sCommerceChat = ""
    SendTxt.Text = ""
End Sub

Private Sub txtAgregar_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 03/10/2009
'**************************************************************
    'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
On Error GoTo ErrHandler
  
    Dim I As Long
    Dim tempstr As String
    Dim CharAscii As Integer
    
    For I = 1 To Len(txtAgregar.text)
        CharAscii = Asc(mid$(txtAgregar.text, I, 1))
        
        If CharAscii >= 48 And CharAscii <= 57 Then
            tempstr = tempstr & Chr$(CharAscii)
        End If
    Next I
    
    If tempstr <> txtAgregar.text Then
        'We only set it if it's different, otherwise the event will be raised
        'constantly and the client will crush
        txtAgregar.text = tempstr
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub txtAgregar_Change de frmComerciarUsu.frm")
End Sub

Private Sub txtAgregar_KeyDown(KeyCode As Integer, Shift As Integer)
If Not ((KeyCode >= 48 And KeyCode <= 57) Or KeyCode = vbKeyBack Or _
        KeyCode = vbKeyDelete Or (KeyCode >= 37 And KeyCode <= 40)) Then
On Error GoTo ErrHandler
  
    KeyCode = 0
End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub txtAgregar_KeyDown de frmComerciarUsu.frm")
End Sub

Private Sub txtAgregar_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or _
        KeyAscii = vbKeyDelete Or (KeyAscii >= 37 And KeyAscii <= 40)) Then
    'txtCant = KeyCode
On Error GoTo ErrHandler
  
    KeyAscii = 0
End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub txtAgregar_KeyPress de frmComerciarUsu.frm")
End Sub

Private Function CheckAvailableSlot(ByVal InvSlot As Byte, ByVal Amount As Long) As Byte
'***************************************************
'Author: ZaMa
'Last Modify Date: 30/11/2009
'Search for an available Slot to put an item. If found returns the Slot, else returns 0.
'***************************************************
    Dim Slot As Long
On Error GoTo Err
    ' Primero chequeo si puedo sumar esa cantidad en algun Slot que ya tenga ese item
    For Slot = 1 To INV_OFFER_SLOTS
        If InvComUsu.ObjIndex(InvSlot) = InvOfferComUsu(0).ObjIndex(Slot) Then
            If InvOfferComUsu(0).Amount(Slot) + Amount <= MAX_INVENTORY_OBJS Then
                ' Puedo sumarlo aca
                CheckAvailableSlot = Slot
                Exit Function
            End If
        End If
    Next Slot
    
    ' No lo puedo sumar, me fijo si hay alguno vacio
    For Slot = 1 To INV_OFFER_SLOTS
        If InvOfferComUsu(0).ObjIndex(Slot) = 0 Then
            ' Esta vacio, lo dejo aca
            CheckAvailableSlot = Slot
            Exit Function
        End If
    Next Slot
    Exit Function
Err:
    Debug.Print "Slot: " & Slot
End Function

Public Sub UpdateInvCom(ByVal ObjIndex As Integer, ByVal Amount As Long)
On Error GoTo ErrHandler
  
    Dim Slot As Byte
    Dim RemainingAmount As Long
    Dim DifAmount As Long
    
    RemainingAmount = Amount
    
    For Slot = 1 To MAX_INVENTORY_SLOTS
        
        If InvComUsu.ObjIndex(Slot) = ObjIndex Then
            DifAmount = Inventario.Amount(Slot) - InvComUsu.Amount(Slot)
            If DifAmount > 0 Then
                If RemainingAmount > DifAmount Then
                    RemainingAmount = RemainingAmount - DifAmount
                    Call InvComUsu.ChangeSlotItemAmount(Slot, Inventario.Amount(Slot))
                Else
                    Call InvComUsu.ChangeSlotItemAmount(Slot, InvComUsu.Amount(Slot) + RemainingAmount)
                    Exit Sub
                End If
            End If
        End If
    Next Slot
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateInvCom de frmComerciarUsu.frm")
End Sub

Public Sub PrintCommerceMsg(ByRef msg As String, ByVal FontIndex As Integer)
On Error GoTo ErrHandler
  
    
    With FontTypes(FontIndex)
        Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, msg, .red, .green, .blue, .bold, .italic)
    End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub PrintCommerceMsg de frmComerciarUsu.frm")
End Sub

Public Function HasAnyItem(ByRef inventory As clsGraphicalInventory) As Boolean
On Error GoTo ErrHandler
  

    Dim Slot As Long
    
    For Slot = 1 To Inventory.MaxObjs
        If Inventory.Amount(Slot) > 0 Then HasAnyItem = True: Exit Function
    Next Slot
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function HasAnyItem de frmComerciarUsu.frm")
End Function

Public Sub HabilitarConfirmar(ByVal Habilitar As Boolean)
On Error GoTo ErrHandler
  
    imgConfirmar.Visible = Habilitar
    Call cBotonConfirmar.EnableButton(Habilitar)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HabilitarConfirmar de frmComerciarUsu.frm")
End Sub

Public Sub HabilitarAceptarRechazar(ByVal Habilitar As Boolean)
On Error GoTo ErrHandler
  
    imgAceptar.Visible = Habilitar
    imgRechazar.Visible = Habilitar
    Call cBotonAceptar.EnableButton(Habilitar)
    Call cBotonRechazar.EnableButton(Habilitar)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub HabilitarAceptarRechazar de frmComerciarUsu.frm")
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CloseWindow
End Sub

Private Sub CloseWindow()
On Error GoTo ErrHandler
    
    Call WriteUserCommerceEnd
    Unload Me
    If frmMain.Visible Then frmMain.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmComerciarUsu.frm")
End Sub

Public Function GetOtherPlayerOfferedGold() As Double
    GetOtherPlayerOfferedGold = OtherPlayerOfferedGold
End Function

Public Sub SetOtherPlayerOfferedGold(ByVal gold As Double)
    OtherPlayerOfferedGold = gold
    
    lblOtherUserGold.Caption = OtherPlayerOfferedGold
End Sub
