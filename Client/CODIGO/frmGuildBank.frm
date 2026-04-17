VERSION 5.00
Begin VB.Form frmGuildBank 
   BorderStyle     =   0  'None
   Caption         =   "Banco de Clan"
   ClientHeight    =   5160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ARGENTUM.AOPictureBox picGMemberInv 
      Height          =   2475
      Left            =   4695
      TabIndex        =   6
      Top             =   1315
      Width           =   2450
      _extentx        =   4313
      _extenty        =   3493
   End
   Begin ARGENTUM.AOPictureBox picGBankInv 
      Height          =   2475
      Left            =   1100
      TabIndex        =   5
      Top             =   1315
      Width           =   2450
      _extentx        =   4313
      _extenty        =   4366
   End
   Begin VB.TextBox TxtQtyGoldBank 
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
      Height          =   220
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   4
      Text            =   "0"
      Top             =   970
      Width           =   750
   End
   Begin VB.TextBox TxtQtyObj 
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
      Height          =   255
      Left            =   3700
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "1"
      Top             =   2070
      Width           =   870
   End
   Begin VB.TextBox TxtQtyGoldUser 
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
      Height          =   220
      Left            =   5760
      MaxLength       =   7
      TabIndex        =   0
      Text            =   "0"
      Top             =   970
      Width           =   750
   End
   Begin VB.Label LblBankGold 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   1430
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label LblUserGold 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   5497
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Image ImgGoldDeposit 
      Height          =   360
      Left            =   5260
      Top             =   930
      Width           =   360
   End
   Begin VB.Image ImgGoldTake 
      Height          =   365
      Left            =   2580
      Top             =   930
      Width           =   365
   End
   Begin VB.Image ImgDeposit 
      Height          =   590
      Left            =   5280
      Top             =   3960
      Width           =   1230
   End
   Begin VB.Image ImgTake 
      Height          =   585
      Left            =   1680
      Top             =   3960
      Width           =   1230
   End
End
Attribute VB_Name = "frmGuildBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents GMemberInv As clsGraphicalInventory
Attribute GMemberInv.VB_VarHelpID = -1
Public WithEvents GBankInv As clsGraphicalInventory
Attribute GBankInv.VB_VarHelpID = -1
Private Const DEFAULTBOX As Integer = 1
 
Private cButtonTakeItem As clsGraphicalButton
Private cButtonDepositItem As clsGraphicalButton
Private cButtonTakeGold As clsGraphicalButton
Private cButtonDepositGold As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CloseWindow
End Sub

Private Sub Form_Load()

    If PlayerData.Guild.IdGuild <= 0 Then Exit Sub

    frmGuildBank.LblBankGold.Caption = PlayerData.Guild.BankGold
    frmGuildBank.LblUserGold.Caption = UserGLD
    
    Call LoadControls
    Call Inicializar
    
    Call modCustomCursors.SetFormCursorDefault(Me)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    If Not GMemberInv Is Nothing Then
        Call GMemberInv.Release
        Set GMemberInv = Nothing
    End If
    
    If Not GBankInv Is Nothing Then
        Call GBankInv.Release
        Set GBankInv = Nothing
    End If
    
End Sub

Public Sub LoadControls()
    Set cButtonTakeItem = New clsGraphicalButton
    Set cButtonDepositItem = New clsGraphicalButton
    Set cButtonTakeGold = New clsGraphicalButton
    Set cButtonDepositGold = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI
    
    Me.Picture = LoadPicture(GrhPath & "VentanaGuildBank.jpg")
    
    Call cButtonTakeItem.Initialize(ImgTake, GrhPath & "BotonGuildBankRetirar.jpg", _
                                    GrhPath & "BotonGuildBankRetirar.jpg", _
                                    GrhPath & "BotonGuildBankRetirar.jpg", Me, _
                                    GrhPath & "BotonGuildBankRetirar_Disabled.jpg") ', _
                                    'GrhPath & "BotonGuildBankRetirar_Click.jpg", Me)
                                    
    Call cButtonDepositItem.Initialize(ImgDeposit, GrhPath & "BotonGuildBankDepositar.jpg", _
                                    GrhPath & "BotonGuildBankDepositar.jpg", _
                                    GrhPath & "BotonGuildBankDepositar.jpg", Me, _
                                    GrhPath & "BotonGuildBankDepositar_Disabled.jpg") ', _
                                    'GrhPath & "BotonGuildBankDepositar_Click.jpg", Me)
                                    
    Call cButtonTakeGold.Initialize(ImgGoldTake, GrhPath & "BotonFlechaDerecha.jpg", _
                                    GrhPath & "BotonFlechaDerecha.jpg", _
                                    GrhPath & "BotonFlechaDerecha.jpg", Me, _
                                    GrhPath & "BotonFlechaDerecha_Disabled.jpg") ', _
                                    'GrhPath & "BotonFlechaDerecha_Click.jpg", Me)
                                    
    Call cButtonDepositGold.Initialize(ImgGoldDeposit, GrhPath & "BotonFlechaIzquierda.jpg", _
                                    GrhPath & "BotonFlechaIzquierda.jpg", _
                                    GrhPath & "BotonFlechaIzquierda.jpg", Me, _
                                    GrhPath & "BotonFlechaIzquierda_Disabled.jpg") ', _
                                    'GrhPath & "BotonFlechaIzquierda_Click.jpg", Me)
    
    
    
End Sub

Private Sub ImgDeposit_Click()
    If Not cButtonDepositItem.IsEnabled Then Exit Sub
    If GMemberInv.SelectedItem = 0 Then Exit Sub
    
    If Not IsNumeric(TxtQtyObj.text) Or TxtQtyObj.text = 0 Then Exit Sub
    
    Call Engine_Audio.PlayInterface(SND_CLICK)
    If GMemberInv.Amount(GMemberInv.SelectedItem) <= 0 Then Exit Sub
    
    Call WriteGuildExchange(eExchangeType.IsObject, eExchangeAction.Deposit, TxtQtyObj.text, GMemberInv.SelectedItem, DEFAULTBOX)
    
    frmGuildMain.GuildBankIsDirty = True
End Sub

Private Sub ImgGoldDeposit_Click()
    If Not cButtonDepositGold.IsEnabled Then Exit Sub
    If Not IsNumeric(TxtQtyGoldUser.text) Or TxtQtyGoldUser.text = 0 Then Exit Sub
    Call Engine_Audio.PlayInterface(SND_CLICK)
    Call WriteGuildExchange(eExchangeType.IsGold, eExchangeAction.Deposit, TxtQtyGoldUser.text, , DEFAULTBOX)
    TxtQtyGoldUser.text = 0
    frmGuildMain.GuildBankIsDirty = True
End Sub

Private Sub ImgGoldTake_Click()
    If Not cButtonTakeGold.IsEnabled Then Exit Sub
    If Not IsNumeric(TxtQtyGoldBank.text) Or TxtQtyGoldBank.text = 0 Then Exit Sub
    Call Engine_Audio.PlayInterface(SND_CLICK)
    Call WriteGuildExchange(eExchangeType.IsGold, eExchangeAction.Withdraw, TxtQtyGoldBank.text, , DEFAULTBOX)
    TxtQtyGoldBank.text = 0
    frmGuildMain.GuildBankIsDirty = True
    
End Sub

Private Sub ImgTake_Click()
    If Not cButtonTakeItem.IsEnabled Then Exit Sub
    If GBankInv.SelectedItem = 0 Then Exit Sub
    
    If Not IsNumeric(TxtQtyObj.text) Or TxtQtyObj.text = 0 Then Exit Sub
    
    Call Engine_Audio.PlayInterface(SND_CLICK)
    If GBankInv.Amount(GBankInv.SelectedItem) <= 0 Then Exit Sub

    Call WriteGuildExchange(eExchangeType.IsObject, eExchangeAction.Withdraw, TxtQtyObj.text, GBankInv.SelectedItem, DEFAULTBOX)
        
    frmGuildMain.GuildBankIsDirty = True
End Sub

Private Sub TxtQtyGoldUser_GotFocus()
    TxtQtyObj.text = 1
End Sub
Private Sub TxtQtyObj_GotFocus()
    TxtQtyGoldUser.text = 0
End Sub
Private Sub TxtQtyGoldUser_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandler

    Call ValKeyPress(KeyAscii)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub TxtQtyGoldUser_KeyPress de frmCantidad.frm")
End Sub



Private Sub TxtQtyObj_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandler
  
    Call ValKeyPress(KeyAscii)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub TxtQtyObj_KeyPress de frmCantidad.frm")
End Sub

Private Sub Inicializar()
On Error GoTo ErrHandler
    Dim I As Integer
    
    If GMemberInv Is Nothing And GBankInv Is Nothing Then
        
        Set GMemberInv = New clsGraphicalInventory
        Set GBankInv = New clsGraphicalInventory
        
        Call GMemberInv.Initialize(frmGuildBank.picGMemberInv, MAX_INVENTORY_SLOTS, , , , 10, , , , , True)
        Call GBankInv.Initialize(frmGuildBank.picGBankInv, PlayerData.Guild.MaxSlotBank, , , , 10, , , , , True)
 
        Set GMemberInv.dropInventory = GBankInv
        Set GBankInv.dropInventory = GMemberInv
        
        Call FillMemberInv
        
        Call FillGuildBankInv
    End If
    
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Inicializar de frmGuildBank.frm")
End Sub


Private Sub ValKeyPress(KeyAscii As Integer)
    Call modHelperFunctions.IsNumericInputKeyPressValid(KeyAscii, True)
    Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Public Sub EnableButtons()
    Call cButtonDepositItem.EnableButton(Guilds.HasPermission(GP_BANK_DEPOSIT_ITEM))
    Call cButtonTakeGold.EnableButton(Guilds.HasPermission(GP_BANK_WITHDRAW_GOLD))
    Call cButtonDepositGold.EnableButton(Guilds.HasPermission(GP_BANK_DEPOSIT_GOLD))
    Call cButtonTakeItem.EnableButton(Guilds.HasPermission(GP_BANK_WITHDRAW_ITEM))
End Sub

Public Sub FillGuildBankInv()
    Dim I As Integer
    ' Fill GuildBank inventory
    For I = 1 To GetQtyGuildBankObjects
            With PlayerData.Guild.Bank(I)
            If .IdObject <> 0 Then
                Call GBankInv.SetItem(I, .IdObject, .Amount, 0, _
                    GameMetadata.Objs(.IdObject).GrhIndex, _
                    GameMetadata.Objs(.IdObject).OBJType, _
                    0, 0, 0, 0, 0, _
                    GameMetadata.Objs(.IdObject).Name, _
                    0, _
                    .CanUse)
            End If
            End With
    Next I
    
    Exit Sub
End Sub

Public Sub FillMemberInv()
    Dim I As Integer
    
    For I = 1 To MAX_INVENTORY_SLOTS
        If Inventario.ObjIndex(I) <> 0 Then
            With Inventario
                Call GMemberInv.SetItem(I, .ObjIndex(I), _
                    .Amount(I), .Equipped(I), .GrhIndex(I), _
                    .OBJType(I), .MaxHit(I), .MinHit(I), .MaxDef(I), .MinDef(I), _
                    .Valor(I), .ItemName(I), 0, .CanUse(I))
            End With
        End If
    Next I

    Exit Sub
End Sub

Public Sub Reload()
    Set GMemberInv = Nothing
    Set GBankInv = Nothing
    Call Form_Load

End Sub
Private Sub CloseWindow()
On Error GoTo ErrHandler
  
    Unload Me
    If frmGuildMain.Visible Then
        Unload frmGuildMain
    End If
    If frmMain.Visible Then frmMain.SetFocus
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmGuildBank.frm")
End Sub
