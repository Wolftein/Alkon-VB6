VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmOpciones 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   12240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   22350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   816
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame framLinks 
      Appearance      =   0  'Flat
      BackColor       =   &H00292929&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   7560
      TabIndex        =   23
      Top             =   7320
      Width           =   6495
      Begin VB.Image imgSupportLinkButton 
         Height          =   525
         Left            =   1920
         Top             =   480
         Width           =   1230
      End
      Begin VB.Image imgManualLinkButton 
         Height          =   525
         Left            =   3240
         Top             =   480
         Width           =   1230
      End
   End
   Begin VB.Frame framExtras 
      Appearance      =   0  'Flat
      BackColor       =   &H00292929&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   14160
      TabIndex        =   22
      Top             =   3720
      Width           =   6495
      Begin VB.Image imgTutorialButton 
         Height          =   525
         Left            =   3240
         Top             =   1200
         Width           =   1230
      End
      Begin VB.Image imgMapButton 
         Height          =   525
         Left            =   1920
         Top             =   1200
         Width           =   1230
      End
      Begin VB.Image imgDialogsButton 
         Height          =   525
         Left            =   4560
         Top             =   480
         Width           =   1230
      End
      Begin VB.Image imgConsoleButton 
         Height          =   525
         Left            =   3240
         Top             =   480
         Width           =   1230
      End
      Begin VB.Image imgMessagesButton 
         Height          =   525
         Left            =   1920
         Top             =   480
         Width           =   1230
      End
      Begin VB.Image imgKeysButton 
         Height          =   525
         Left            =   600
         Top             =   480
         Width           =   1230
      End
   End
   Begin VB.Frame framInterface 
      Appearance      =   0  'Flat
      BackColor       =   &H00292929&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   7560
      TabIndex        =   14
      Top             =   3720
      Width           =   6495
      Begin VB.ComboBox cboCursorSets 
         Height          =   315
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2120
         Width           =   1695
      End
      Begin VB.TextBox txtGuildNewsQty 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   5610
         MaxLength       =   1
         TabIndex        =   21
         Text            =   "9"
         Top             =   1800
         Width           =   270
      End
      Begin VB.ComboBox cboGuildDialog 
         Height          =   315
         Left            =   4180
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1395
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Usar paquete de cursores"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   27
         Top             =   2160
         Width           =   3495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de Líneas de Diálogos"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   20
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Image imgCheckShowLatency 
         Height          =   225
         Left            =   5640
         Top             =   720
         Width           =   210
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Ver Latencia"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   720
         Width           =   4935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Ver FPS"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Ver Noticias de Clan"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   16
         Top             =   1080
         Width           =   4935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ver Diálogos de Clanes"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   1440
         Width           =   4695
      End
      Begin VB.Image imgCheckShowGuildNews 
         Height          =   225
         Left            =   5640
         Top             =   1080
         Width           =   210
      End
      Begin VB.Image imgCheckShowFPS 
         Height          =   225
         Left            =   5640
         Top             =   360
         Width           =   210
      End
   End
   Begin VB.Frame framGraphics 
      Appearance      =   0  'Flat
      BackColor       =   &H00292929&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   14160
      TabIndex        =   10
      Top             =   120
      Width           =   6495
      Begin VB.Image imgVerticalSync 
         Height          =   225
         Left            =   5640
         Top             =   1080
         Width           =   210
      End
      Begin VB.Image imgCheckAskForResolutionChange 
         Height          =   225
         Left            =   5640
         Top             =   720
         Width           =   210
      End
      Begin VB.Image imgCheckFullScreen 
         Height          =   225
         Left            =   5640
         Top             =   360
         Width           =   210
      End
      Begin VB.Label lblFullScreen 
         BackStyle       =   0  'Transparent
         Caption         =   "Pantalla Completa"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblAskForResolutionChange 
         BackStyle       =   0  'Transparent
         Caption         =   "Preguntar por cambio de resolución"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   720
         Width           =   4935
      End
      Begin VB.Label lblCompatibilityMode 
         BackStyle       =   0  'Transparent
         Caption         =   "Vertical Sync"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   1080
         Width           =   4935
      End
   End
   Begin VB.Frame framAudio 
      Appearance      =   0  'Flat
      BackColor       =   &H00292929&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   7560
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.TextBox txtInterfaceSoundValue 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   5640
         MaxLength       =   3
         TabIndex        =   26
         Text            =   "100"
         Top             =   1440
         Width           =   435
      End
      Begin VB.TextBox txtEffectSoundValue 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   5640
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "100"
         Top             =   1080
         Width           =   435
      End
      Begin VB.TextBox txtMusicValue 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   5640
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "100"
         Top             =   720
         Width           =   435
      End
      Begin VB.TextBox txtMasterSoundValue 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   5640
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "100"
         Top             =   360
         Width           =   435
      End
      Begin MSComctlLib.Slider sldMasterSound 
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         Top             =   360
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider sldMusic 
         Height          =   255
         Left            =   2760
         TabIndex        =   5
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider sldEffectSound 
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider sldInterfaceSound 
         Height          =   255
         Left            =   2760
         TabIndex        =   25
         Top             =   1440
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         TickStyle       =   3
      End
      Begin VB.Image imgCheckEnableInterfaceSound 
         Height          =   225
         Left            =   2400
         Top             =   1440
         Width           =   210
      End
      Begin VB.Label lblInterfaceSound 
         BackStyle       =   0  'Transparent
         Caption         =   "Volumen de interfaz"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   24
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Image imgCheckEnableEffectSound 
         Height          =   225
         Left            =   2400
         Top             =   1090
         Width           =   210
      End
      Begin VB.Image imgCheckEnableMusic 
         Height          =   225
         Left            =   2400
         Top             =   730
         Width           =   210
      End
      Begin VB.Image imgCheckEnableMasterSound 
         Height          =   225
         Left            =   2400
         Top             =   370
         Width           =   210
      End
      Begin VB.Label lblEffectSound 
         BackStyle       =   0  'Transparent
         Caption         =   "Volumen de efectos"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblMusic 
         BackStyle       =   0  'Transparent
         Caption         =   "Volumen de música"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblMasterSound 
         BackStyle       =   0  'Transparent
         Caption         =   "Volumen principal"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Image imgAcceptButton 
      Height          =   540
      Left            =   3120
      Top             =   6750
      Width           =   1230
   End
   Begin VB.Image imgLinksButton 
      Height          =   540
      Left            =   5850
      Top             =   900
      Width           =   1230
   End
   Begin VB.Image imgExtrasButton 
      Height          =   540
      Left            =   4470
      Top             =   900
      Width           =   1230
   End
   Begin VB.Image imgInterfaceButton 
      Height          =   540
      Left            =   3120
      Top             =   900
      Width           =   1230
   End
   Begin VB.Image imgGraphicsButton 
      Height          =   540
      Left            =   1770
      Top             =   900
      Width           =   1230
   End
   Begin VB.Image imgAudioButton 
      Height          =   540
      Left            =   420
      Top             =   900
      Width           =   1230
   End
End
Attribute VB_Name = "frmOpciones"
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

Private cButtonConfigKeys As clsGraphicalButton
Private cButtonPersonalizedMsg As clsGraphicalButton
Private cButtonMap As clsGraphicalButton
Private cButtonManual As clsGraphicalButton
Private cButtonSupport As clsGraphicalButton
Private cButtonTutorial As clsGraphicalButton
Private cButtonConsole As clsGraphicalButton
Private cButtonDialogs As clsGraphicalButton
Private cButtonAccept As clsGraphicalButton

Private cButtonAudioMenu As clsGraphicalButton
Private cButtonGraphicsMenu As clsGraphicalButton
Private cButtonInterfaceMenu As clsGraphicalButton
Private cButtonExtrasMenu As clsGraphicalButton
Private cButtonLinksMenu As clsGraphicalButton

Private cButtonEnableMasterSound As clsGraphicalButton
Private cButtonEnableMusic As clsGraphicalButton
Private cButtonEnableEffectSound As clsGraphicalButton
Private cButtonEnableInterfaceSound As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private PicCheckBoxEnabled As Picture
Private picCheckBoxDisabled As Picture

Private Loading As Boolean

Private Const CONTAINER_POSITION_LEFT As Byte = 32
Private Const CONTAINER_POSITION_TOP As Byte = 120

Private CurrentFrame As Frame

Private Sub SetControlSizes()
    Me.Height = 7500
    Me.Width = 7500
End Sub

Private Sub ShowFrame(ByRef Control As Frame)

    If Not CurrentFrame Is Nothing Then
        If Control.Name = CurrentFrame.Name Then Exit Sub
    End If
    
    Control.Visible = True

    Control.Left = CONTAINER_POSITION_LEFT
    Control.Top = CONTAINER_POSITION_TOP
    
    Control.Width = 433
    Control.Height = 233
    
    Call SetControlSizes
        
    If Not CurrentFrame Is Nothing Then
        If CurrentFrame.Name <> "" Then
             CurrentFrame.Visible = False
        End If
    End If

    Set CurrentFrame = Control
End Sub


Private Sub cboCursorSets_Click()
    GameConfig.Extras.MouseCursorSetToUse = CByte(cboCursorSets.ItemData(cboCursorSets.ListIndex))
End Sub


Private Sub cboGuildDialog_Click()
    GameConfig.Guilds.bShowDialogsInConsole = IIf(cboGuildDialog.ListIndex = 1, True, False)
    
    txtGuildNewsQty.Enabled = Not GameConfig.Guilds.bShowDialogsInConsole
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
On Error GoTo ErrHandler
   
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "VentanaOpciones.jpg")
    Call LoadButtons
    
    Loading = True      'Prevent sounds when setting check's values
    
    Call LoadCursorSetsComboOptions
    Call LoadGuildDialogComboOptions
    
    Call SetControlSizes
    
    ' Load the default frame.
    Call imgAudioButton_Click
    
    Loading = False     'Enable sounds when setting check's values
    
    Call modCustomCursors.SetFormCursorDefault(Me)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Form_Load de frmOpciones.frm")
End Sub

Private Sub LoadCursorSetsComboOptions()
    Call cboCursorSets.Clear
    
    Call cboCursorSets.AddItem("Alkon", 0)
    cboCursorSets.ItemData(0) = eMousePointerPackType.Default
    
    Call cboCursorSets.AddItem("Custom", 1)
    cboCursorSets.ItemData(1) = eMousePointerPackType.Custom
    
End Sub

Private Sub LoadGuildDialogComboOptions()
    Call cboGuildDialog.Clear
    
    Call cboGuildDialog.AddItem("En Pantalla", 0)
    Call cboGuildDialog.AddItem("En Consola", 1)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub ImgAcceptButton_Click()

    Set CurrentFrame = Nothing
    
    Call mod_Configuration.SaveGameConfig
    Call CerrarVentana
End Sub

Private Sub imgAudioButton_Click()
    Call ShowFrame(framAudio)
    
    Call DrawCheck(imgCheckEnableMasterSound, GameConfig.Sounds.bMasterSoundEnabled)
    Call DrawCheck(imgCheckEnableMusic, GameConfig.Sounds.bMusicEnabled)
    Call DrawCheck(imgCheckEnableEffectSound, GameConfig.Sounds.bSoundEffectsEnabled)
    Call DrawCheck(imgCheckEnableInterfaceSound, GameConfig.Sounds.bInterfaceEnabled)
    
    Call DrawSlider(sldMasterSound, GameConfig.Sounds.MasterVolume, GameConfig.Sounds.bMasterSoundEnabled)
    Call DrawSlider(sldMusic, GameConfig.Sounds.MusicVolume, GameConfig.Sounds.bMusicEnabled)
    Call DrawSlider(sldEffectSound, GameConfig.Sounds.SoundsVolume, GameConfig.Sounds.bSoundEffectsEnabled)
    Call DrawSlider(sldInterfaceSound, GameConfig.Sounds.InterfaceVolume, GameConfig.Sounds.bInterfaceEnabled)
    
    txtMasterSoundValue.Enabled = GameConfig.Sounds.bMasterSoundEnabled
    txtMusicValue.Enabled = GameConfig.Sounds.bMusicEnabled
    txtEffectSoundValue.Enabled = GameConfig.Sounds.bSoundEffectsEnabled
    txtInterfaceSoundValue.Enabled = GameConfig.Sounds.bInterfaceEnabled
    
    txtMasterSoundValue.text = CStr(GameConfig.Sounds.MasterVolume)
    txtMusicValue.text = CStr(GameConfig.Sounds.MusicVolume)
    txtEffectSoundValue.text = CStr(GameConfig.Sounds.SoundsVolume)
    txtInterfaceSoundValue.text = CStr(GameConfig.Sounds.InterfaceVolume)
    
End Sub

Private Sub DrawCheck(ByRef ImageControl As Image, ByVal Property As Boolean)
    ImageControl.Picture = IIf(Property = True, PicCheckBoxEnabled, picCheckBoxDisabled)
End Sub

Private Sub DrawCombobox(ByRef ComboboxControl As ComboBox, ByVal value As Byte)
    ComboboxControl.ListIndex = value
End Sub

Private Sub DrawSlider(ByRef SliderControl As Slider, ByVal value As Integer, ByVal EnabledControl As Boolean)
    If value > SliderControl.Max Then value = SliderControl.Max
    
    SliderControl.value = value
    SliderControl.Enabled = EnabledControl
End Sub

Private Sub imgCheckAskForResolutionChange_Click()
On Error GoTo ErrHandler
        
    GameConfig.Extras.bAskForResolutionChange = Not GameConfig.Extras.bAskForResolutionChange
       
    Call DrawCheck(imgCheckAskForResolutionChange, GameConfig.Extras.bAskForResolutionChange)

  Exit Sub

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgCheckAskForResolutionChange_Click de frmOpciones.frm")
End Sub

Private Sub imgCheckEnableInterfaceSound_Click()
        
    GameConfig.Sounds.bInterfaceEnabled = Not GameConfig.Sounds.bInterfaceEnabled
    
    Call DrawCheck(imgCheckEnableInterfaceSound, GameConfig.Sounds.bInterfaceEnabled)
        
    Engine_Audio.InterfaceEnabled = GameConfig.Sounds.bInterfaceEnabled
    
    sldInterfaceSound.Enabled = GameConfig.Sounds.bInterfaceEnabled
    txtInterfaceSoundValue.Enabled = GameConfig.Sounds.bInterfaceEnabled
End Sub

Private Sub imgCheckEnableEffectSound_Click()
        
    GameConfig.Sounds.bSoundEffectsEnabled = Not GameConfig.Sounds.bSoundEffectsEnabled
    
    Call DrawCheck(imgCheckEnableEffectSound, GameConfig.Sounds.bSoundEffectsEnabled)
        
    Engine_Audio.EffectEnabled = GameConfig.Sounds.bSoundEffectsEnabled
    
    sldEffectSound.Enabled = GameConfig.Sounds.bSoundEffectsEnabled
    txtEffectSoundValue.Enabled = GameConfig.Sounds.bSoundEffectsEnabled
End Sub

Private Sub imgCheckEnableMasterSound_Click()

    GameConfig.Sounds.bMasterSoundEnabled = Not GameConfig.Sounds.bMasterSoundEnabled
    
    Call DrawCheck(imgCheckEnableMasterSound, GameConfig.Sounds.bMasterSoundEnabled)
            
    Engine_Audio.MasterEnabled = GameConfig.Sounds.bMasterSoundEnabled
    
    sldMasterSound.Enabled = GameConfig.Sounds.bMasterSoundEnabled
    txtMasterSoundValue.Enabled = GameConfig.Sounds.bMasterSoundEnabled
End Sub

Private Sub imgCheckEnableMusic_Click()
        
    GameConfig.Sounds.bMusicEnabled = Not GameConfig.Sounds.bMusicEnabled
    
    Call DrawCheck(imgCheckEnableMusic, GameConfig.Sounds.bMusicEnabled)
    
    Engine_Audio.MusicEnabled = GameConfig.Sounds.bMusicEnabled
    
    sldMusic.Enabled = GameConfig.Sounds.bMusicEnabled
    txtMusicValue.Enabled = GameConfig.Sounds.bMusicEnabled
End Sub

Private Sub imgCheckShowFPS_Click()
        
    GameConfig.Extras.bShowFPS = Not GameConfig.Extras.bShowFPS
    
    Call DrawCheck(imgCheckShowFPS, GameConfig.Extras.bShowFPS)
End Sub

Private Sub imgCheckShowGuildNews_Click()
    GameConfig.Guilds.bShowGuildNews = Not GameConfig.Guilds.bShowGuildNews
    
    Call DrawCheck(imgCheckShowGuildNews, GameConfig.Guilds.bShowGuildNews)
End Sub

Private Sub imgCheckShowLatency_Click()
        
    GameConfig.Extras.bShowLatency = Not GameConfig.Extras.bShowLatency
    
    Call DrawCheck(imgCheckShowLatency, GameConfig.Extras.bShowLatency)
End Sub

Private Sub imgConsoleButton_Click()
        
    Call frmConfigMsg.Show(vbModeless, Me)
End Sub

Private Sub imgDialogsButton_Click()
        
    Call frmDialogos.Show(vbModeless, Me)
End Sub

Private Sub imgExtrasButton_Click()
    Call ShowFrame(framExtras)
End Sub

Private Sub imgInterfaceButton_Click()
    Call ShowFrame(framInterface)
    
    Call DrawCheck(imgCheckShowGuildNews, GameConfig.Guilds.bShowGuildNews)
    Call DrawCheck(imgCheckShowFPS, GameConfig.Extras.bShowFPS)
    Call DrawCheck(imgCheckShowLatency, GameConfig.Extras.bShowLatency)
    Call DrawCombobox(cboGuildDialog, IIf(GameConfig.Guilds.bShowDialogsInConsole, 1, 0))
    Call SelectItemCboMousePointers(GameConfig.Extras.MouseCursorSetToUse)
    txtGuildNewsQty.text = GameConfig.Guilds.MaxMessageQuantity
End Sub

Private Sub imgKeysButton_Click()

    Call frmCustomKeys.Show(vbModeless, Me)
End Sub

Private Sub imgLinksButton_Click()
    Call ShowFrame(framLinks)
End Sub

Private Sub imgManualLinkButton_Click()
On Error GoTo ErrHandler
  
    Call ShellExecute(0, "Open", "https://manual.alkononline.com.ar", "", App.path, SW_SHOWNORMAL)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgManualLinkButton_Click de frmOpciones.frm")
End Sub

Private Sub imgMapButton_Click()

    Call ShellExecute(0, "Open", MAP_URL, "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub imgMessagesButton_Click()
 
    Call frmMessageTxt.Show(vbModeless, Me)
End Sub

Private Sub imgSupportLinkButton_Click()
On Error GoTo ErrHandler
  
    Call ShellExecute(0, "Open", "https://www.alkononline.com.ar", "", App.path, SW_SHOWNORMAL)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgSupportLinkButton_Click de frmOpciones.frm")
End Sub

Private Sub imgTutorialButton_Click()
        
    Call frmTutorial.Show(vbModeless, Me)
End Sub

Private Sub imgVerticalSync_Click()
On Error GoTo ErrHandler

    If Loading Then Exit Sub
    
    GameConfig.Graphics.bUseVerticalSync = Not GameConfig.Graphics.bUseVerticalSync
       
    Call DrawCheck(imgVerticalSync, GameConfig.Graphics.bUseVerticalSync)

  Exit Sub

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgUseCompatibilityMode_Click de frmOpciones.frm")
End Sub

Private Sub imgCheckFullScreen_Click()
On Error GoTo ErrHandler

    If Loading Then Exit Sub
    
    GameConfig.Graphics.bUseFullScreen = Not GameConfig.Graphics.bUseFullScreen
       
    Call DrawCheck(imgCheckFullScreen, GameConfig.Graphics.bUseFullScreen)

  Exit Sub

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub imgCheckFullScreen_Click de frmOpciones.frm")
End Sub

Private Sub imgGraphicsButton_Click()
    Call ShowFrame(framGraphics)
    
    Call DrawCheck(imgCheckAskForResolutionChange, GameConfig.Extras.bAskForResolutionChange)
    Call DrawCheck(imgVerticalSync, GameConfig.Graphics.bUseVerticalSync)
    Call DrawCheck(imgCheckFullScreen, GameConfig.Graphics.bUseFullScreen)
    
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
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CerrarVentana de frmOpciones.frm")
End Sub

Private Sub LoadButtons()
On Error GoTo ErrHandler
  
    Dim GrhPath As String
    
    GrhPath = DirInterfaces & SELECTED_UI
    
    Set cButtonConfigKeys = New clsGraphicalButton
    Set cButtonPersonalizedMsg = New clsGraphicalButton
    Set cButtonMap = New clsGraphicalButton
    Set cButtonManual = New clsGraphicalButton
    Set cButtonSupport = New clsGraphicalButton
    Set cButtonTutorial = New clsGraphicalButton
    Set cButtonDialogs = New clsGraphicalButton
    Set cButtonConsole = New clsGraphicalButton
    Set cButtonAccept = New clsGraphicalButton
    
    Set cButtonAudioMenu = New clsGraphicalButton
    Set cButtonGraphicsMenu = New clsGraphicalButton
    Set cButtonInterfaceMenu = New clsGraphicalButton
    Set cButtonExtrasMenu = New clsGraphicalButton
    Set cButtonLinksMenu = New clsGraphicalButton
    
    Set cButtonEnableMasterSound = New clsGraphicalButton
    Set cButtonEnableMusic = New clsGraphicalButton
    Set cButtonEnableEffectSound = New clsGraphicalButton
    Set cButtonEnableInterfaceSound = New clsGraphicalButton

    Set LastButtonPressed = New clsGraphicalButton
    
    Call cButtonConfigKeys.Initialize(imgKeysButton, GrhPath & "BotonTeclas.jpg", _
                                    GrhPath & "BotonTeclas.jpg", _
                                    GrhPath & "BotonTeclas.jpg", Me)
                                    
    Call cButtonPersonalizedMsg.Initialize(imgMessagesButton, GrhPath & "BotonMensajes.jpg", _
                                    GrhPath & "BotonMensajes.jpg", _
                                    GrhPath & "BotonMensajes.jpg", Me)
                                    
    Call cButtonMap.Initialize(imgMapButton, GrhPath & "BotonMapaOpciones.jpg", _
                                    GrhPath & "BotonMapaOpciones.jpg", _
                                    GrhPath & "BotonMapaOpciones.jpg", Me)
                                    
    Call cButtonManual.Initialize(imgManualLinkButton, GrhPath & "BotonManual.jpg", _
                                    GrhPath & "BotonManual.jpg", _
                                    GrhPath & "BotonManual.jpg", Me)
                                                      
    Call cButtonSupport.Initialize(imgSupportLinkButton, GrhPath & "BotonSoporte.jpg", _
                                    GrhPath & "BotonSoporte.jpg", _
                                    GrhPath & "BotonSoporte.jpg", Me)
                                    
    Call cButtonTutorial.Initialize(imgTutorialButton, GrhPath & "BotonTutorial.jpg", _
                                    GrhPath & "BotonTutorial.jpg", _
                                    GrhPath & "BotonTutorial.jpg", Me)
                                    
    Call cButtonDialogs.Initialize(imgDialogsButton, GrhPath & "BotonDialogos.jpg", _
                                    GrhPath & "BotonDialogos.jpg", _
                                    GrhPath & "BotonDialogos.jpg", Me)
                                 
    Call cButtonConsole.Initialize(imgConsoleButton, GrhPath & "BotonConsola.jpg", _
                                    GrhPath & "BotonConsola.jpg", _
                                    GrhPath & "BotonConsola.jpg", Me)
                                    
    Call cButtonAccept.Initialize(imgAcceptButton, GrhPath & "BotonAceptar.jpg", _
                                    GrhPath & "BotonAceptar.jpg", _
                                    GrhPath & "BotonAceptar.jpg", Me)
                       
    Call cButtonAudioMenu.Initialize(imgAudioButton, GrhPath & "BotonAudioMenu.jpg", _
                                    GrhPath & "BotonAudioMenu.jpg", _
                                    GrhPath & "BotonAudioMenu.jpg", Me)
                                    
    Call cButtonGraphicsMenu.Initialize(imgGraphicsButton, GrhPath & "BotonGraficosMenu.jpg", _
                                    GrhPath & "BotonGraficosMenu.jpg", _
                                    GrhPath & "BotonGraficosMenu.jpg", Me)
                                    
    Call cButtonInterfaceMenu.Initialize(imgInterfaceButton, GrhPath & "BotonInterfazMenu.jpg", _
                                    GrhPath & "BotonInterfazMenu.jpg", _
                                    GrhPath & "BotonInterfazMenu.jpg", Me)
                                    
    Call cButtonExtrasMenu.Initialize(imgExtrasButton, GrhPath & "BotonExtrasMenu.jpg", _
                                    GrhPath & "BotonExtrasMenu.jpg", _
                                    GrhPath & "BotonExtrasMenu.jpg", Me)
                                    
    Call cButtonLinksMenu.Initialize(imgLinksButton, GrhPath & "BotonLinksMenu.jpg", _
                                    GrhPath & "BotonLinksMenu.jpg", _
                                    GrhPath & "BotonLinksMenu.jpg", Me)
                                    
                         
    Set PicCheckBoxEnabled = LoadPicture(GrhPath & "CheckEnabled.jpg")
    Set picCheckBoxDisabled = LoadPicture(GrhPath & "CheckDisabled.jpg")
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadButtons de frmOpciones.frm")
End Sub

Private Sub sldMasterSound_Scroll()
On Error GoTo ErrHandler

    GameConfig.Sounds.MasterVolume = sldMasterSound.value
    Engine_Audio.MasterVolume = GameConfig.Sounds.MasterVolume
    txtMasterSoundValue.text = CStr(GameConfig.Sounds.MasterVolume)
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub sldMasterSound_Scroll de frmOpciones.frm")
End Sub

Private Sub sldInterfaceSound_Scroll()
On Error GoTo ErrHandler

    GameConfig.Sounds.InterfaceVolume = sldInterfaceSound.value
    Engine_Audio.InterfaceVolume = GameConfig.Sounds.InterfaceVolume
    txtInterfaceSoundValue.text = GameConfig.Sounds.InterfaceVolume
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub sldInterfaceSound_Scroll de frmOpciones.frm")
End Sub

Private Sub sldEffectSound_Scroll()
On Error GoTo ErrHandler

    GameConfig.Sounds.SoundsVolume = sldEffectSound.value
    Engine_Audio.EffectVolume = GameConfig.Sounds.SoundsVolume
    txtEffectSoundValue.text = GameConfig.Sounds.SoundsVolume
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub sldEffectSound_Scroll de frmOpciones.frm")
End Sub

Private Sub sldMusic_Scroll()
On Error GoTo ErrHandler

    GameConfig.Sounds.MusicVolume = sldMusic.value
    Engine_Audio.MusicVolume = GameConfig.Sounds.MusicVolume
    txtMusicValue.text = GameConfig.Sounds.MusicVolume
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub sldMusic_Scroll de frmOpciones.frm")
End Sub

Private Sub txtGuildNewsQty_Change()
    On Error GoTo ErrHandler
  
    txtGuildNewsQty.text = Val(txtGuildNewsQty.text)
    
    DialogosClanes.CantidadDialogos = txtGuildNewsQty.text
    GameConfig.Guilds.MaxMessageQuantity = DialogosClanes.CantidadDialogos
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub txtGuildNewsQty_Change de frmOpciones.frm")
End Sub

Private Sub CheckKeyAsciiNumber(ByRef KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or _
        KeyAscii = vbKeyDelete Or (KeyAscii >= 37 And KeyAscii <= 40)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtInterfaceSoundValue_Change()
On Error GoTo ErrHandler

    If Val(txtInterfaceSoundValue.text) > 100 Then
        txtInterfaceSoundValue.text = 100
    End If
    
    GameConfig.Sounds.InterfaceVolume = Val(txtInterfaceSoundValue.text)
    sldInterfaceSound.value = GameConfig.Sounds.InterfaceVolume
    txtInterfaceSoundValue.text = GameConfig.Sounds.InterfaceVolume
    
    If GameConfig.Sounds.bInterfaceEnabled Then
        Engine_Audio.InterfaceVolume = GameConfig.Sounds.InterfaceVolume
    End If
    
    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub txtInterfaceSoundValue_Change de frmOpciones.frm")
End Sub

Private Sub txtEffectSoundValue_Change()
On Error GoTo ErrHandler

    If Val(txtEffectSoundValue.text) > 100 Then
        txtEffectSoundValue.text = 100
    End If
    
    GameConfig.Sounds.SoundsVolume = Val(txtEffectSoundValue.text)
    
    sldEffectSound.value = GameConfig.Sounds.SoundsVolume
    txtEffectSoundValue.text = GameConfig.Sounds.SoundsVolume
    
    If GameConfig.Sounds.bSoundEffectsEnabled Then
        Engine_Audio.EffectVolume = GameConfig.Sounds.SoundsVolume
    End If
    
    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub txtEffectSoundValue_Change de frmOpciones.frm")
End Sub

Private Sub txtMusicValue_Change()
On Error GoTo ErrHandler
    
    If Val(txtMusicValue.text) > 100 Then
        txtMusicValue.text = 100
    End If

    GameConfig.Sounds.MusicVolume = Val(txtMusicValue.text)
    sldMusic.value = GameConfig.Sounds.MusicVolume
    txtMusicValue.text = GameConfig.Sounds.MusicVolume
    
    If GameConfig.Sounds.bMusicEnabled Then
        Engine_Audio.MusicVolume = GameConfig.Sounds.MusicVolume
    End If
    
    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub txtMusicValue_Change de frmOpciones.frm")
End Sub

Private Sub txtMasterSoundValue_Change()
On Error GoTo ErrHandler

    If Val(txtMasterSoundValue.text) > 100 Then
        txtMasterSoundValue.text = 100
    End If
   
    GameConfig.Sounds.MasterVolume = Val(txtMasterSoundValue.text)
    sldMasterSound.value = GameConfig.Sounds.MasterVolume
    txtMasterSoundValue.text = GameConfig.Sounds.MasterVolume
    
    If GameConfig.Sounds.bMasterSoundEnabled Then
        Engine_Audio.MasterVolume = GameConfig.Sounds.MasterVolume
    End If
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub txtMasterSoundValue_Change de frmOpciones.frm")
End Sub

Private Sub txtInterfaceSoundValue_KeyPress(KeyAscii As Integer)
    Call CheckKeyAsciiNumber(KeyAscii)
End Sub

Private Sub txtEffectSoundValue_KeyPress(KeyAscii As Integer)
    Call CheckKeyAsciiNumber(KeyAscii)
End Sub

Private Sub txtMasterSoundValue_KeyPress(KeyAscii As Integer)
    Call CheckKeyAsciiNumber(KeyAscii)
End Sub

Private Sub txtMusicValue_KeyPress(KeyAscii As Integer)
    Call CheckKeyAsciiNumber(KeyAscii)
End Sub


Private Sub SelectItemCboMousePointers(ByVal MousePointerpack As Byte)
    Dim FoundElement As Byte
    FoundElement = 0
    
    Dim I As Integer
    
    For I = 0 To cboCursorSets.ListCount - 1
        If CByte(cboCursorSets.ItemData(I)) = MousePointerpack Then
            cboCursorSets.ListIndex = I
            Exit Sub
        End If
    Next I
    
    cboCursorSets.ListIndex = 0
    
End Sub

