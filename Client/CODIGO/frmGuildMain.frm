VERSION 5.00
Begin VB.Form frmGuildMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Clan"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   489.022
   ScaleMode       =   0  'User
   ScaleWidth      =   750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PboContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00292929&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4938
      Left            =   2955
      ScaleHeight     =   4935
      ScaleWidth      =   7920
      TabIndex        =   0
      Top             =   1641
      Width           =   7920
   End
   Begin VB.Image ImgBlockBank 
      Height          =   840
      Left            =   480
      ToolTipText     =   "Debe Comprar el upgrade"
      Top             =   3480
      Width           =   990
   End
   Begin VB.Image ImgClose 
      Height          =   525
      Left            =   6480
      Top             =   6750
      Width           =   1215
   End
   Begin VB.Image ImgGuildSettings 
      Height          =   855
      Left            =   960
      Top             =   6360
      Width           =   975
   End
   Begin VB.Image ImgRelations 
      Height          =   855
      Left            =   1560
      Top             =   5429
      Width           =   975
   End
   Begin VB.Image ImgTrophies 
      Height          =   855
      Left            =   480
      Top             =   5429
      Width           =   975
   End
   Begin VB.Image ImgQuests 
      Height          =   855
      Left            =   1530
      Top             =   4478
      Width           =   975
   End
   Begin VB.Image ImgUpgrades 
      Height          =   855
      Left            =   480
      Top             =   4478
      Width           =   975
   End
   Begin VB.Image ImgBank 
      Height          =   855
      Left            =   480
      Top             =   3527
      Width           =   2055
   End
   Begin VB.Image ImgMembers 
      Height          =   855
      Left            =   480
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Image ImgInfo 
      Height          =   855
      Left            =   480
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblGuildName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre clan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
End
Attribute VB_Name = "frmGuildMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ActiveForm As Form
Public GuildBankIsDirty As Boolean

Private cButtonInfo As clsGraphicalButton
Private cButtonMembers As clsGraphicalButton
Private cButtonBank As clsGraphicalButton
Private cButtonUpgrades As clsGraphicalButton
Private cButtonMissions As clsGraphicalButton
Private cButtonTrophies As clsGraphicalButton
Private cButtonRelations As clsGraphicalButton
Private cButtonSettings As clsGraphicalButton
Private cButtonBlockBank As clsGraphicalButton

Private cButtonClose As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CloseWindow
End Sub

Private Sub Form_Load()
    lblGuildName.Caption = PlayerData.Guild.Name
    GuildBankIsDirty = False
    
    Call LoadControls
    
    Call modCustomCursors.SetFormCursorDefault(Me)
End Sub

Public Sub LoadForm(ByVal frm As Form, ByVal Title As String)
    If Not ActiveForm Is Nothing Then Call ActiveForm.Hide
    
    Set ActiveForm = frm
    Call frm.Show
    'lblActiveFormLabel.Caption = Title
    SetParent frm.hwnd, PboContainer.hwnd
    frm.Move 0, 0, PboContainer.ScaleWidth, PboContainer.ScaleHeight
    frm.Visible = True
End Sub

Private Sub LoadControls()
    
    Set cButtonInfo = New clsGraphicalButton
    Set cButtonMembers = New clsGraphicalButton
    Set cButtonBank = New clsGraphicalButton
    Set cButtonUpgrades = New clsGraphicalButton
    Set cButtonMissions = New clsGraphicalButton
    Set cButtonTrophies = New clsGraphicalButton
    Set cButtonRelations = New clsGraphicalButton
    Set cButtonSettings = New clsGraphicalButton
    Set cButtonClose = New clsGraphicalButton
    Set cButtonBlockBank = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI
    
    Me.Picture = LoadPicture(GrhPath & "VentanaGuildMain.jpg")
    
    Call cButtonInfo.Initialize(ImgInfo, GrhPath & "BotonGuildMainInfo.jpg", _
                                    GrhPath & "BotonGuildMainInfo.jpg", _
                                    GrhPath & "BotonGuildMainInfo.jpg", Me)
                                    
    Call cButtonMembers.Initialize(ImgMembers, GrhPath & "BotonGuildMainMembers.jpg", _
                                    GrhPath & "BotonGuildMainMembers.jpg", _
                                    GrhPath & "BotonGuildMainMembers.jpg", Me)
                                    
    Call cButtonBank.Initialize(ImgBank, GrhPath & "BotonGuildMainBank.jpg", _
                                    GrhPath & "BotonGuildMainBank.jpg", _
                                    GrhPath & "BotonGuildMainBank.jpg", Me, _
                                    GrhPath & "BotonGuildMainBank_Disabled.jpg")
                                    
    Call cButtonUpgrades.Initialize(ImgUpgrades, GrhPath & "BotonGuildMainUpgrades.jpg", _
                                    GrhPath & "BotonGuildMainUpgrades.jpg", _
                                    GrhPath & "BotonGuildMainUpgrades.jpg", Me, _
                                    GrhPath & "BotonGuildMainUpgrades_Disabled.jpg")

    Call cButtonMissions.Initialize(ImgQuests, GrhPath & "BotonGuildMainQuests.jpg", _
                                    GrhPath & "BotonGuildMainQuests.jpg", _
                                    GrhPath & "BotonGuildMainQuests.jpg", Me, _
                                    GrhPath & "BotonGuildMainQuests_Disabled.jpg")
                                    
    Call cButtonTrophies.Initialize(ImgTrophies, GrhPath & "BotonGuildMainTrophies.jpg", _
                                    GrhPath & "BotonGuildMainTrophies.jpg", _
                                    GrhPath & "BotonGuildMainTrophies.jpg", Me, _
                                    GrhPath & "BotonGuildMainTrophies_Disabled.jpg")
                                    
    Call cButtonRelations.Initialize(ImgRelations, GrhPath & "BotonGuildMainRelations.jpg", _
                                    GrhPath & "BotonGuildMainRelations.jpg", _
                                    GrhPath & "BotonGuildMainRelations.jpg", Me, _
                                    GrhPath & "BotonGuildMainRelations_Disabled.jpg")
                                    
    Call cButtonSettings.Initialize(ImgGuildSettings, GrhPath & "BotonGuildMainSettings.jpg", _
                                    GrhPath & "BotonGuildMainSettings.jpg", _
                                    GrhPath & "BotonGuildMainSettings.jpg", Me)
                                    
                                    
    Call cButtonClose.Initialize(ImgClose, GrhPath & "BotonGuildMainClose.jpg", _
                                    GrhPath & "BotonGuildMainClose.jpg", _
                                    GrhPath & "BotonGuildMainClose.jpg", Me)
                                   
    Call cButtonBlockBank.Initialize(ImgBlockBank, GrhPath & "BotonGuildBlockBank.jpg", _
                                    GrhPath & "BotonGuildBlockBank.jpg", _
                                    GrhPath & "BotonGuildBlockBank.jpg", Me)
                                                                  
    If PlayerData.Guild.BankAvalaible Then
        ImgBlockBank.Visible = False
    Else
        ImgBlockBank.Visible = True
    End If
    
    ' Disabling some options by default as they're not going to be available for now.
    Call cButtonTrophies.EnableButton(False)
    Call cButtonRelations.EnableButton(False)
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set PboContainer = Nothing
    
    Unload frmGuildBank
    Unload frmGuildInformation
    Unload frmGuildEditRoles
    Unload frmGuildMembers
    Unload frmGuildQuests
    Unload frmGuildRelations
    Unload frmGuildTrophies
    Unload frmGuildUpgrades
    Unload frmGuildRoleAssign
    Unload frmGuildMemberAdd
    Unload frmGuildRolesPermissions
    Unload frmGuildQuestActive
    
End Sub

Public Sub CloseForm()
    Call ActiveForm.Hide
    Call Me.Hide
    If GuildBankIsDirty Then
        Call WriteGuildBankEnd
        GuildBankIsDirty = False
    End If
    Call LoadForm(frmGuildInformation, "Información del clan")
    If frmMain.Visible Then frmMain.SetFocus
End Sub

Private Sub ImgClose_Click()
    Call CloseForm
End Sub

Private Sub ImgBank_Click()
    If Not cButtonBank.IsEnabled Then Exit Sub
    
    Call LoadForm(frmGuildBank, "Banco")
    Call frmGuildBank.EnableButtons
    
    Call frmGuildBank.FillMemberInv
    Call frmGuildBank.FillGuildBankInv
End Sub

Private Sub ImgGuildSettings_Click()
    Call frmGuildRolesList.UpdateRoleList
    Call LoadForm(frmGuildRolesList, "Edición de Roles")
    Call frmGuildRolesList.CleanUCs
    Call frmGuildRolesList.InitializeUCs
End Sub

Private Sub ImgInfo_Click()
    Call LoadForm(frmGuildInformation, "Información del clan")
End Sub

Private Sub ImgMembers_Click()
    Call LoadForm(frmGuildMembers, "Miembros")
    Call frmGuildMembers.CleanUCs
    Call frmGuildMembers.InitializeUCs
    Call frmGuildMembers.UpdateMemberQty
End Sub

Private Sub ImgQuests_Click()
    If Not cButtonMissions.IsEnabled Then Exit Sub
    
    If PlayerData.Guild.Quest.Id = 0 Then
        Call LoadForm(frmGuildQuests, "Misiones")
        Call frmGuildQuests.ShowData
    Else
        Call LoadForm(frmGuildQuestActive, "Misiones")
        Call frmGuildQuestActive.ShowData
    End If
End Sub

Private Sub ImgRelations_Click()
    If Not cButtonRelations.IsEnabled Then Exit Sub
    
    Call LoadForm(frmGuildRelations, "Relaciones")
End Sub

Private Sub ImgTrophies_Click()
    If Not cButtonTrophies.IsEnabled Then Exit Sub
    
    Call LoadForm(frmGuildTrophies, "Trofeos")
End Sub

Private Sub ImgUpgrades_Click()
    If Not cButtonUpgrades.IsEnabled Then Exit Sub
    
    Call LoadForm(frmGuildUpgrades, "Mejoras")
End Sub

Public Sub ShowPartial()
    Me.Show , frmMain
    Call LoadForm(frmGuildInformation, "Información del clan")
    Call DisableOptions
    PlayerData.Guild.IsFullFormOpen = False
        
    If PlayerData.Guild.BankAvalaible Then
        Call cButtonBank.EnableButton(True)
    Else
        Call cButtonBank.EnableButton(False)
    End If
End Sub

Public Sub ShowFull()
    Me.Show , frmMain
    Call LoadForm(frmGuildInformation, "Información del clan")
    Call DisableOptions
    PlayerData.Guild.IsFullFormOpen = True
    If PlayerData.Guild.BankAvalaible Then
        Call cButtonBank.EnableButton(True)
    Else
        Call cButtonBank.EnableButton(False)
    End If
End Sub

Public Sub DisableOptions()
    If (Guilds.HasPermission(GP_BANK_DEPOSIT_GOLD) Or _
        Guilds.HasPermission(GP_BANK_DEPOSIT_ITEM) Or _
        Guilds.HasPermission(GP_BANK_WITHDRAW_GOLD) Or _
        Guilds.HasPermission(GP_BANK_WITHDRAW_ITEM)) And _
        PlayerData.Guild.IsFullFormOpen Then
        Call cButtonBank.EnableButton(True)
    Else
        Call cButtonBank.EnableButton(False)
        PlayerData.Guild.IsFullFormOpen = False
    End If
End Sub
Private Sub CloseWindow()
On Error GoTo ErrHandler
  
    Unload Me
    If frmMain.Visible Then frmMain.SetFocus
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWindow de frmGuildMain.frm")
End Sub
