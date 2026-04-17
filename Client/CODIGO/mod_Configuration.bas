Attribute VB_Name = "mod_Configuration"
Option Explicit

Public Const GFX_PATH As String = "\Graficos\"
Public Const SND_PATH As String = "\Wav\"
Public Const MSC_PATH As String = "\MP3\"
Public Const MAP_PATH As String = "\Mapas\"

Private Const INIT_PATH As String = "\INIT\"

Public Type tGraphicsAmbienceConfig
    bLightsEnabled As Boolean
    bAmbientLightsEnabled As Boolean
    bUseRainWithParticles As Boolean
End Type
    
Public Type tGraphicsConfig
    Ambience As tGraphicsAmbienceConfig
    bUseFullScreen As Boolean '!bNoRes
    bUseVerticalSync As Boolean
End Type
    
Public Type tGuildsConfig
    bShowGuildNews As Boolean       'bGuildNews
    bShowDialogsInConsole As Boolean 'bGldMsgConsole
    MaxMessageQuantity As Byte      'bCantMsgs
End Type
    
Public Type tSoundsConfig
    bMasterSoundEnabled As Boolean
    MasterVolume As Byte
    bMusicEnabled As Boolean        '!bNoMusic
    MusicVolume As Byte
    bSoundEffectsEnabled As Boolean '!bNoSoundEffects
    SoundsVolume As Byte
    bInterfaceEnabled As Boolean    '!bNoSoundEffects
    InterfaceVolume As Byte
End Type
    
Public Type tExtraConfig
    Name As String
    NameStyle As Byte               ' Nombres
    bRightClickEnabled As Boolean   'rightClickActivated
    bAskForResolutionChange As Boolean
    bShowLatency As Boolean
    bShowFPS As Boolean
    MouseCursorSetToUse As Byte
End Type
    
Public Type tGameConfig
    Graphics    As tGraphicsConfig
    Sounds      As tSoundsConfig
    Guilds      As tGuildsConfig
    Extras      As tExtraConfig
End Type
    
Public GameConfig As tGameConfig

Public Sub LoadGameConfig()
    On Error GoTo ErrHandler
    Dim Path As String
    Dim Parser As TOMLParser
    Set Parser = New TOMLParser
    
    Dim vbData As String
    Dim vbFile As Long
    vbFile = FreeFile
    Path = App.Path & "/INIT/UserConfig.ini"
    
    If FileExist(Path, vbNormal) Then
        Open App.Path + "/INIT/UserConfig.ini" For Input Access Read As vbFile
            vbData = Input$(LOF(vbFile), vbFile)
        Close vbFile
    End If

    Call Parser.Load(vbData)
    
    Call LoadExtrasConfig(Parser.GetSection("EXTRAS"))
    Call LoadGraphicsConfig(Parser.GetSection("GRAPHICSENGINE"))
    Call LoadSoundsConfig(Parser.GetSection("SOUND"))
    Call LoadGuildConfig(Parser.GetSection("GUILD"))

    Call SaveGameConfig
    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadGameConfig de mod_Configuration.bas")
End Sub
    
Public Sub SaveGameConfig()
    On Error GoTo ErrHandler
    Dim Path As String
    Dim Parser As TOMLParser
    Set Parser = New TOMLParser

    Call SaveExtrasConfig(Parser.GetSection("EXTRAS"))
    Call SaveGraphicsConfig(Parser.GetSection("GRAPHICSENGINE"))
    Call SaveSoundsConfig(Parser.GetSection("SOUND"))
    Call SaveGuildConfig(Parser.GetSection("GUILD"))
    
    Dim vbFile As Long
    vbFile = FreeFile
    
    Path = App.Path & "/INIT/UserConfig.ini"
    
    If FileExist(Path, vbNormal) Then Kill Path
    Open Path For Binary Access Write As vbFile
        Put vbFile, , Parser.Dump()
    Close vbFile

    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SaveGameConfig de mod_Configuration.bas")
End Sub

Private Sub LoadSoundsConfig(ByVal Section As TOMLSection)
    GameConfig.Sounds.bMasterSoundEnabled = Section.GetBool("MasterEnabled", True)
    GameConfig.Sounds.bMusicEnabled = Section.GetBool("MusicEnabled", True)
    GameConfig.Sounds.bSoundEffectsEnabled = Section.GetBool("SoundEffectsEnabled", True)
    GameConfig.Sounds.bInterfaceEnabled = Section.GetInt8("InterfaceEnabled", True)
    GameConfig.Sounds.MasterVolume = Section.GetInt8("MasterVolume", 50)
    GameConfig.Sounds.MusicVolume = Section.GetInt8("MusicVolume", 50)
    GameConfig.Sounds.SoundsVolume = Section.GetInt8("SoundsVolume", 50)
    GameConfig.Sounds.InterfaceVolume = Section.GetInt8("InterfaceVolume", 50)
End Sub
    
Private Sub LoadExtrasConfig(ByVal Section As TOMLSection)
    GameConfig.Extras.Name = Section.GetString("Name", vbNullString)
    GameConfig.Extras.NameStyle = Section.GetInt8("NameStyle", 2)
    GameConfig.Extras.bRightClickEnabled = Section.GetBool("RightClickEnabled", True)
    GameConfig.Extras.bAskForResolutionChange = Section.GetBool("AskForResolutionChange", True)
    GameConfig.Extras.MouseCursorSetToUse = Section.GetInt8("MouseCursorSetToUse", 1)
    GameConfig.Extras.bShowFPS = Section.GetBool("ShowFps", False)
    GameConfig.Extras.bShowLatency = Section.GetBool("ShowLatency", False)
End Sub
    
Private Sub LoadGuildConfig(ByVal Section As TOMLSection)
    GameConfig.Guilds.bShowDialogsInConsole = Section.GetBool("ShowDialogsInConsole", True)
    GameConfig.Guilds.bShowGuildNews = Section.GetBool("ShowGuildNews", True)
    GameConfig.Guilds.MaxMessageQuantity = Section.GetInt8("MaxMessageQuantity", 5)
End Sub
        
Private Sub LoadGraphicsConfig(ByVal Section As TOMLSection)
    GameConfig.Graphics.bUseFullScreen = Section.GetBool("UseFullScreen", False)
    GameConfig.Graphics.bUseVerticalSync = Section.GetBool("UseVerticalSync", False)
    GameConfig.Graphics.Ambience.bAmbientLightsEnabled = Section.GetBool("EnableAmbientLights", True)
    GameConfig.Graphics.Ambience.bLightsEnabled = Section.GetBool("EnableLights", True)
    GameConfig.Graphics.Ambience.bUseRainWithParticles = Section.GetBool("UseRainWithParticles", False)
End Sub
    
Private Sub SaveGraphicsConfig(ByVal Section As TOMLSection)
    Call Section.SetBool("UseFullScreen", GameConfig.Graphics.bUseFullScreen)
    Call Section.SetBool("UseVerticalSync", GameConfig.Graphics.bUseVerticalSync)
    Call Section.SetBool("EnableAmbientLights", GameConfig.Graphics.Ambience.bAmbientLightsEnabled)
    Call Section.SetBool("EnableLights", GameConfig.Graphics.Ambience.bLightsEnabled)
    Call Section.SetBool("UseRainWithParticles", GameConfig.Graphics.Ambience.bUseRainWithParticles)
End Sub
    
Private Sub SaveExtrasConfig(ByVal Section As TOMLSection)
    Call Section.SetString("Name", GameConfig.Extras.Name)
    Call Section.SetInt8("NameStyle", GameConfig.Extras.NameStyle)
    Call Section.SetBool("RightClickEnabled", GameConfig.Extras.bRightClickEnabled)
    Call Section.SetBool("AskForResolutionChange", GameConfig.Extras.bAskForResolutionChange)
    Call Section.SetInt8("MouseCursorSetToUse", GameConfig.Extras.MouseCursorSetToUse)
    Call Section.SetBool("ShowFps", GameConfig.Extras.bShowFPS)
    Call Section.SetBool("ShowLatency", GameConfig.Extras.bShowLatency)
End Sub
    
Private Sub SaveSoundsConfig(ByVal Section As TOMLSection)
    Call Section.SetBool("MasterEnabled", GameConfig.Sounds.bMasterSoundEnabled)
    Call Section.SetBool("MusicEnabled", GameConfig.Sounds.bMusicEnabled)
    Call Section.SetBool("SoundEffectsEnabled", GameConfig.Sounds.bSoundEffectsEnabled)
    Call Section.SetBool("InterfaceEnabled", GameConfig.Sounds.bInterfaceEnabled)
    Call Section.SetInt8("MasterVolume", GameConfig.Sounds.MasterVolume)
    Call Section.SetInt8("MusicVolume", GameConfig.Sounds.MusicVolume)
    Call Section.SetInt8("SoundsVolume", GameConfig.Sounds.SoundsVolume)
    Call Section.SetInt8("InterfaceVolume", GameConfig.Sounds.InterfaceVolume)
End Sub
    
Private Sub SaveGuildConfig(ByVal Section As TOMLSection)
    Call Section.SetBool("ShowDialogsInConsole", GameConfig.Guilds.bShowDialogsInConsole)
    Call Section.SetBool("ShowGuildNews", GameConfig.Guilds.bShowGuildNews)
    Call Section.SetInt8("MaxMessageQuantity", GameConfig.Guilds.MaxMessageQuantity)
End Sub
      
    
