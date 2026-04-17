VERSION 5.00
Begin VB.Form frmBancoObj 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6915
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   530
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   461
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ARGENTUM.AOPictureBox PicBancoInv 
      Height          =   3345
      Left            =   540
      TabIndex        =   5
      Top             =   1710
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   5900
   End
   Begin VB.TextBox CantidadOro 
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
      Height          =   210
      Left            =   3240
      MaxLength       =   7
      TabIndex        =   4
      Text            =   "1"
      Top             =   1905
      Width           =   555
   End
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   210
      Left            =   3240
      MaxLength       =   5
      TabIndex        =   3
      Text            =   "1"
      Top             =   3390
      Width           =   555
   End
   Begin ARGENTUM.AOPictureBox PicInv 
      Height          =   3345
      Left            =   4020
      TabIndex        =   6
      Top             =   1710
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   5900
   End
   Begin VB.Image imgGoToAccountBank 
      Height          =   525
      Left            =   5400
      Tag             =   "0"
      Top             =   270
      Width           =   1230
   End
   Begin VB.Label lblUserGld 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1500
      TabIndex        =   2
      Top             =   1335
      Width           =   1455
   End
   Begin VB.Image imgDepositarOro 
      Height          =   270
      Left            =   3210
      Tag             =   "0"
      Top             =   1515
      Width           =   270
   End
   Begin VB.Image imgRetirarOro 
      Height          =   270
      Left            =   3555
      Tag             =   "0"
      Top             =   1515
      Width           =   270
   End
   Begin VB.Image imgCerrar 
      Height          =   525
      Left            =   2850
      Tag             =   "0"
      Top             =   7140
      Width           =   1230
   End
   Begin VB.Image imgDepositItem 
      Height          =   270
      Left            =   3210
      MousePointer    =   99  'Custom
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image imgWhitdrawItem 
      Height          =   270
      Left            =   3555
      MousePointer    =   99  'Custom
      Top             =   3000
      Width           =   270
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ItemDescription"
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
      Height          =   915
      Index           =   1
      Left            =   1995
      TabIndex        =   1
      Top             =   5880
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NombreItem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   5595
      Visible         =   0   'False
      Width           =   3060
   End
End
Attribute VB_Name = "frmBancoObj"
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

'[CODE]:MatuX
'
'    Le puse el iconito de la manito a los botones ^_^ y
'   le puse borde a la ventana.
'
'[END]'

'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->

Private clsFormulario As clsFormMovementManager

Private cBotonRetirarOro As clsGraphicalButton
Private cBotonDepositarOro As clsGraphicalButton

Private cButtonWhitdrawItem As clsGraphicalButton
Private cButtonDepositItem As clsGraphicalButton

Private cButtonGoToBankAccount As clsGraphicalButton

Private cBotonCerrar As clsGraphicalButton

' WithEvents not allowed in arrays.
Private WithEvents dragBankInventory As clsGraphicalInventory
Attribute dragBankInventory.VB_VarHelpID = -1
Private WithEvents dragUserInventory As clsGraphicalInventory
Attribute dragUserInventory.VB_VarHelpID = -1

Public LastButtonPressed As clsGraphicalButton

Public LasActionBuy As Boolean
Public LastIndex1 As Integer
Public LastIndex2 As Integer
Public NoPuedeMover As Boolean

Private Sub dragBankInventory_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer, _
                                       ByVal moveType As eMoveType)
On Error GoTo ErrHandler
  
    Select Case moveType
        
        Case eMoveType.InventoryToTarget
            imgFlecha_Click (1)
    End Select
    
    Me.Refresh
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub dragBankInventory_dragDone de frmBancoObj.frm")
End Sub

Private Sub dragBankInventory_updateSlotInfo(ByVal nSlot As Integer)
    ' Current slot.
On Error GoTo ErrHandler
  
    Call PicBancoInv_Click
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub dragBankInventory_updateSlotInfo de frmBancoObj.frm")
End Sub

Private Sub dragUserInventory_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer, _
                                       ByVal moveType As eMoveType)
On Error GoTo ErrHandler
  
    Select Case moveType
        
        Case eMoveType.InventoryToTarget
            imgFlecha_Click (0)
    End Select
    
    Me.Refresh
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub dragUserInventory_dragDone de frmBancoObj.frm")
End Sub

Private Sub dragUserInventory_updateSlotInfo(ByVal nSlot As Integer)
    ' Current slot.
On Error GoTo ErrHandler
  
    Call PicInv_Click
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub dragUserInventory_updateSlotInfo de frmBancoObj.frm")
End Sub

Private Sub cantidad_Change()

On Error GoTo ErrHandler
  
    If Val(cantidad.text) < 1 Then
        cantidad.text = 1
    End If
    
    If Val(cantidad.text) > MAX_INVENTORY_OBJS Then
        cantidad.text = MAX_INVENTORY_OBJS
    End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub cantidad_Change de frmBancoObj.frm")
End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
On Error GoTo ErrHandler
  
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub cantidad_KeyPress de frmBancoObj.frm")
End Sub

Private Sub CantidadOro_Change()
    If Val(CantidadOro.text) < 1 Then
On Error GoTo ErrHandler
  
        cantidad.text = 1
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CantidadOro_Change de frmBancoObj.frm")
End Sub

Private Sub CantidadOro_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
On Error GoTo ErrHandler
  
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CantidadOro_KeyPress de frmBancoObj.frm")
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
On Error GoTo ErrHandler
  
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me, , False
    
    ' Inventories.
    Set dragBankInventory = InvBanco(1)
    Set dragUserInventory = InvBanco(0)
    
    If Not dragBankInventory Is Nothing Then Set dragUserInventory.dropInventory = dragBankInventory
    If Not dragUserInventory Is Nothing Then Set dragBankInventory.dropInventory = dragUserInventory
    
    Call LoadControls
    
    Call modCustomCursors.SetFormCursorDefault(Me)

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmBancoObj.frm")
End Sub

Private Sub LoadControls()
On Error GoTo ErrHandler
  
    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI

    Me.Picture = LoadPicture(GrhPath & "VentanaBoveda.jpg")
        
    Set cBotonRetirarOro = New clsGraphicalButton
    Set cBotonDepositarOro = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    
    Set cButtonDepositItem = New clsGraphicalButton
    Set cButtonWhitdrawItem = New clsGraphicalButton
    
    Set cButtonGoToBankAccount = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton



    Call cButtonDepositItem.Initialize(ImgDepositItem, GrhPath & "BotonFlechaIzquierda_2.jpg", GrhPath & "BotonFlechaIzquierda_2.jpg", GrhPath & "BotonFlechaIzquierda_2.jpg", Me)
    Call cButtonWhitdrawItem.Initialize(imgWhitdrawItem, GrhPath & "BotonFlechaDerecha_2.jpg", GrhPath & "BotonFlechaDerecha_2.jpg", GrhPath & "BotonFlechaDerecha_2.jpg", Me)
    
    Call cButtonGoToBankAccount.Initialize(imgGoToAccountBank, GrhPath & "BotonCuentaFlechaDerecha.jpg", GrhPath & "BotonCuentaFlechaDerecha.jpg", GrhPath & "BotonCuentaFlechaDerecha.jpg", Me)
    

    Call cBotonDepositarOro.Initialize(imgDepositarOro, GrhPath & "BotonFlechaIzquierda_2.jpg", GrhPath & "BotonFlechaIzquierda_2.jpg", GrhPath & "BotonFlechaIzquierda_2.jpg", Me)
    Call cBotonRetirarOro.Initialize(imgRetirarOro, GrhPath & "BotonFlechaDerecha_2.jpg", GrhPath & "BotonFlechaDerecha_2.jpg", GrhPath & "BotonFlechaDerecha_2.jpg", Me)
    

    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCerrar.jpg", GrhPath & "BotonCerrar.jpg", GrhPath & "BotonCerrar.jpg", Me)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadButtons de frmBancoObj.frm")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LastButtonPressed.ToggleToNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandler
  
    Set dragBankInventory = Nothing
    Set dragUserInventory = Nothing
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Unload de frmBancoObj.frm")
End Sub

Private Sub imgFlecha_Click(Index As Integer)
    
On Error GoTo ErrHandler
  
    Call Engine_Audio.PlayInterface(SND_CLICK)
    
    If InvBanco(Index).SelectedItem = 0 Then Exit Sub
    
    If Not IsNumeric(cantidad.text) Then Exit Sub
    
    Select Case Index
        Case 0
            LastIndex1 = InvBanco(0).SelectedItem
            LasActionBuy = True
            Call WriteBankExtractItem(InvBanco(0).SelectedItem, cantidad.text)
            
       Case 1
            LastIndex2 = InvBanco(1).SelectedItem
            LasActionBuy = False
            Call WriteBankDeposit(InvBanco(1).SelectedItem, cantidad.text)
    End Select

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgFlecha_Click de frmBancoObj.frm")
End Sub


Private Sub imgDepositarOro_Click()
    Call WriteBankDepositGold(Val(CantidadOro.text))
End Sub

Private Sub ImgDepositItem_Click()
On Error GoTo ErrHandler
  
    Call Engine_Audio.PlayInterface(SND_CLICK)
    
    If InvBanco(1).SelectedItem = 0 Then Exit Sub
    If Not IsNumeric(cantidad.text) Then Exit Sub

    LastIndex2 = InvBanco(1).SelectedItem
    LasActionBuy = False
    Call WriteBankDeposit(InvBanco(1).SelectedItem, cantidad.text)

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgDepositItem_Click de frmBancoObj.frm")
End Sub

Private Sub imgGoToAccountBank_Click()
    Call CerrarVentana
    Call WriteAccBankStart(vbNullString)
End Sub

Private Sub imgRetirarOro_Click()
    Call WriteBankExtractGold(Val(CantidadOro.text))
End Sub

Private Sub imgWhitdrawItem_Click()
On Error GoTo ErrHandler
  
    Call Engine_Audio.PlayInterface(SND_CLICK)
    
    If InvBanco(0).SelectedItem = 0 Then Exit Sub
    If Not IsNumeric(cantidad.text) Then Exit Sub

    LastIndex2 = InvBanco(0).SelectedItem
    LasActionBuy = False
    Call WriteBankExtractItem(InvBanco(0).SelectedItem, cantidad.text)

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgWhitdrawItem_Click de frmBancoObj.frm")
End Sub

Private Sub PicBancoInv_Click()
On Error GoTo ErrHandler
  
    If dragBankInventory.Dragging Then Exit Sub
    Call InvBanco(1).DeselectItem
    
    If InvBanco(0).SelectedItem <= 0 Then
        Label1(0).Caption = ""
        Label1(1).Visible = False
        Exit Sub
    End If
    
    
    Dim ItemMinLevel As Byte
    Dim ShowItemLevel As Boolean
    Dim ItemDescription As String
    
    With UserBancoInventory(InvBanco(0).SelectedItem)
        Label1(0).Caption = .Name
        Label1(0).Visible = True
        
        ItemMinLevel = GameMetadata.Objs(.ObjIndex).MinimumLevel
        
        Select Case .OBJType
            Case 2, 32
                ItemDescription = "Golpe Mínimo: " & .MinHit & vbCrLf & "Golpe Máximo: " & .MaxHit
                ShowItemLevel = True
                Label1(1).Visible = True
                
            Case 3, 16, 17
                ItemDescription = "Defensa Mín: " & .MinDef & vbCrLf & "Defensa Máx: " & .MaxDef
                ShowItemLevel = True
                Label1(1).Visible = True
                
            Case Else
                Label1(1).Caption = ""
                Label1(1).Visible = False
                Exit Sub
                                    
        End Select
        
        If ShowItemLevel And ItemMinLevel > 0 Then
            ItemDescription = ItemDescription & vbCrLf & "Nivel Mínimo: " & ItemMinLevel
        End If
        
        Label1(1).Caption = ItemDescription
        Label1(1).Visible = True
        
    End With

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub PicBancoInv_Click de frmBancoObj.frm")
End Sub

Private Sub PicBancoInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LastButtonPressed.ToggleToNormal
End Sub

Private Sub PicInv_Click()
On Error GoTo ErrHandler
  
    If dragUserInventory.Dragging Then Exit Sub
    Call InvBanco(0).DeselectItem
    
    If InvBanco(1).SelectedItem <= 0 Or InvBanco(1).SelectedItem > InvBanco(1).MaxObjs Then
        Label1(0).Caption = ""
        Label1(1).Visible = False
        Exit Sub
    End If
    
    Dim ItemMinLevel As Byte
    Dim ShowItemLevel As Boolean
    Dim ItemDescription As String
    
    With Inventario
        Label1(0).Caption = .ItemName(InvBanco(1).SelectedItem)
        Label1(0).Visible = True
        
        ItemMinLevel = GameMetadata.Objs(.ObjIndex(InvBanco(1).SelectedItem)).MinimumLevel
        
        Select Case .OBJType(InvBanco(1).SelectedItem)
            Case eObjType.otWeapon, eObjType.otFlechas, eObjType.otTool
                ItemDescription = "Golpe Mínimo: " & .MinHit(InvBanco(1).SelectedItem) & vbCrLf & "Golpe Máximo: " & .MaxHit(InvBanco(1).SelectedItem)
                ShowItemLevel = True
                Label1(1).Visible = True
                
            Case eObjType.otCasco, eObjType.otArmadura, eObjType.otEscudo ' 3, 16, 17
                ItemDescription = "Defensa Mínima: " & .MinDef(InvBanco(1).SelectedItem) & vbCrLf & "Defensa Máxima: " & .MaxDef(InvBanco(1).SelectedItem)
                ShowItemLevel = True
            Case Else
                ShowItemLevel = False
                Label1(1).Caption = ""
                Label1(1).Visible = False
                Exit Sub
                
        End Select
        
        If ShowItemLevel And ItemMinLevel > 0 Then
            ItemDescription = ItemDescription & vbCrLf & "Nivel Mínimo: " & ItemMinLevel
        End If
        
        Label1(1).Caption = ItemDescription
        Label1(1).Visible = True
        
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub PicInv_Click de frmBancoObj.frm")
End Sub

Private Sub PicInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgCerrar_Click()
    Call CerrarVentana
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CerrarVentana
End Sub

Private Sub CerrarVentana()
On Error GoTo ErrHandler
  
    Call dragBankInventory.Release
    Call dragUserInventory.Release
    
    Set dragBankInventory = Nothing
    Set dragUserInventory = Nothing
    
    Call WriteBankEnd
    NoPuedeMover = False
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CerrarVentana de frmBancoObj.frm")
End Sub
