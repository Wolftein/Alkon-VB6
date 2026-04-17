Attribute VB_Name = "Mod_TileEngine"
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
'@Folder("TileEngine")
Option Explicit

Public Const ENGINE_SPEED As Single = 0.018

Public ExtendRender As clsExtendRenderBase

''''''''''''''''''''''''''''''''''''''''''
'Map sizes in tiles
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

Public Const GrhFogata As Integer = 1521

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1

'Posicion en un mapa
Public Type Position
    X As Long
    Y As Long
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
Public Type GrhData
    sX As Integer
    sY As Integer
    
    fileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    S0 As Single
    T0 As Single
    S1 As Single
    T1 As Single
    
    Speed As Single
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Single
    FrameTimer As Single
    Speed As Single
    Started As Byte
    Loops As Integer
    Alpha As Single
    Depth As Single
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Fx
Public Type CharFx
    FxGrh As Grh
    FxIndex As Integer
End Type

'Apariencia del personaje
Public Type Char
    active As Byte
    Heading As E_Heading
    Pos As Position
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    
    Alignment As eCharacterAlignment
    
    Fx(0 To 10) As CharFx
    LastFx As Byte
    
    glowfX As Grh
    glowFxIndex As Integer
    glowEnabled As Boolean
    
    Criminal As Byte
    Atacable As Boolean
    
    Nombre As String
    
    
    OverheadIcon As Integer
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    UseInvisibilityAlpha As Boolean
    priv As Byte
    'overHeadAnimation As IAnimation

    auraIndex As Integer
    ' NPC attributes
    bHostile As Boolean
    bMerchant As Boolean
    NpcNumber As Integer
    IsSailing As Boolean
    
    LastSpellCast As Integer
    
    ' Sound 3D (Cool)
    SoundSource As Audio_Emitter
    
    Node As Partitioner_Item
End Type

'Info de un objeto
Public Type Obj
    ObjNode As Partitioner_Item
    ObjIndex As Integer
    Amount As Integer
    Luminous As Boolean
    LightOffsetX As Integer
    LightOffsetY As Integer
    LightSize As Integer
    CanBeTransparent As Boolean
    SoundSource As Audio_Emitter ' Objects can make sounds too!
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    Nodes(1 To 4)   As Partitioner_Item
    
    CharIndex As Integer
    ObjGrh As Grh
    
    NpcIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
End Type

Public Type tFuente
    Asset As Graphic_Font
    Tamanio As Long
    color As Long
End Type

Public Type tFuentesJuego
    FuenteBase As tFuente
    
    'Nicks
    NickNewbie As tFuente
    NickNeutral As tFuente
    NickFactionRoyal As tFuente
    NickFactionLegion As tFuente
   
    NickConcilio As tFuente
    NickConsejo As tFuente
    NickDios As tFuente
    NickSemidios As tFuente
    NickConsejero As tFuente
    NickAdmins As tFuente
    NickRolemasters As tFuente
    NickNpcs As tFuente
    
    'General
    Talk As tFuente
    Fight As tFuente
    Warning As tFuente
    Info As tFuente
    InfoBold As tFuente
    Execution As tFuente
    Party As tFuente
    Poison As tFuente
    Guild As tFuente
    Server As tFuente
    GuildMsg As tFuente
    GMSG As tFuente
    
    ConsejoVesA As tFuente
    ConcilioVesA As tFuente
    
    Inventarios As tFuente

End Type

Public FuentesJuego As tFuentesJuego

Public IniPath As String
Public MapPath As String


'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Status del user
Public CurMap As Integer 'Mapa actual
Public UserIndex As Integer
Public WaitInput As Boolean
Public UserMoving As Boolean
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public EngineRun As Boolean

Public FPS As Long
Public FramesPerSecCounter As Long
Private fpsLastCheck As Long


Public HalfWindowTileWidth As Integer
Public HalfWindowTileHeight As Integer

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Private timerEngine        As Currency
Private timerElapsedTime   As Single
Private timerTicksPerFrame As Single

Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer

Private MouseTileX As Byte
Private MouseTileY As Byte

#If EnableSecurity Then

Public MI(1 To 1233) As clsManagerInvisibles
Public CualMI As Integer

#End If


'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData() As MapBlock ' Mapa
Public MapInfo As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public bRain        As Boolean 'está raineando?
Public bTecho       As Boolean 'hay techo?

Public charlist(1 To 10000) As Char

'[CODE 001]:MatuX
Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum
'[END]'
'
'       [END]
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public MinLimiteX As Integer
Public MaxLimiteX As Integer
Public MinLimiteY As Integer
Public MaxLimiteY As Integer

Private g_Technique_1 As Graphic_Pipeline
Private g_Technique_2 As Graphic_Pipeline

Private g_Rain_Material  As Graphic_Material
Public g_Last_OffsetX    As Single
Public g_Last_OffsetY    As Single

'''''''''''''''''''''''''''''''''''''''''''''''

Public Const CHAT_DEPTH As Byte = 5
Private Const TEXT_OUTLINE_COLOR = &HFF212121


''''''' Alpha ''''''''''''

'Techos
Private Const ROOF_ALPHA_SPEED = 0.3
Private Const ROOF_ALPHA_MAX As Byte = 200
Private Const ROOF_ALPHA_MIN As Byte = 0
Private RoofAlpha As Single

'Arboles
Private Const OBJECT_ALPHA_SPEED = 0.25
Private Const OBJECT_ALPHA_MAX As Byte = 255
Private Const OBJECT_ALPHA_MIN = 127

Private CurrentUIDevice As Long
Private CurrentUIDeviceCamera As Graphic_Camera

Sub CargarCabezas()
On Error GoTo ErrHandler
  
    Dim N As Integer
    Dim I As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open App.path & "\init\Cabezas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For I = 1 To Numheads
        Get #N, , Miscabezas(I)
        
        If Miscabezas(I).Head(1) Then
            Call InitGrh(HeadData(I).Head(1), Miscabezas(I).Head(1), 0)
            Call InitGrh(HeadData(I).Head(2), Miscabezas(I).Head(2), 0)
            Call InitGrh(HeadData(I).Head(3), Miscabezas(I).Head(3), 0)
            Call InitGrh(HeadData(I).Head(4), Miscabezas(I).Head(4), 0)
        End If
    Next I
    
    Close #N
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CargarCabezas de TileEngine.bas")
End Sub

Sub CargarCascos()
On Error GoTo ErrHandler
  
    Dim N As Integer
    Dim I As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open App.path & "\init\Cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For I = 1 To NumCascos
        Get #N, , Miscabezas(I)
        
        If Miscabezas(I).Head(1) Then
            Call InitGrh(CascoAnimData(I).Head(1), Miscabezas(I).Head(1), 0)
            Call InitGrh(CascoAnimData(I).Head(2), Miscabezas(I).Head(2), 0)
            Call InitGrh(CascoAnimData(I).Head(3), Miscabezas(I).Head(3), 0)
            Call InitGrh(CascoAnimData(I).Head(4), Miscabezas(I).Head(4), 0)
        End If
    Next I
    
    Close #N
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CargarCascos de TileEngine.bas")
End Sub

Sub CargarCuerpos()
On Error GoTo ErrHandler
  
    Dim N As Integer
    Dim I As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile()
    Open App.path & "\init\Personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For I = 1 To NumCuerpos
        Get #N, , MisCuerpos(I)
        
        If MisCuerpos(I).Body(1) Then
            InitGrh BodyData(I).Walk(1), MisCuerpos(I).Body(1), 0
            InitGrh BodyData(I).Walk(2), MisCuerpos(I).Body(2), 0
            InitGrh BodyData(I).Walk(3), MisCuerpos(I).Body(3), 0
            InitGrh BodyData(I).Walk(4), MisCuerpos(I).Body(4), 0
            
            BodyData(I).HeadOffset.X = MisCuerpos(I).HeadOffsetX
            BodyData(I).HeadOffset.Y = MisCuerpos(I).HeadOffsetY
        End If
    Next I
    
    Close #N
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CargarCuerpos de TileEngine.bas")
End Sub

Sub CargarFxs()
On Error GoTo ErrHandler
  
    Dim N As Integer
    Dim I As Long
    Dim NumFxs As Integer
    
    N = FreeFile()
    Open App.path & "\init\Fxs.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For I = 1 To NumFxs
        Get #N, , FxData(I)
    Next I
    
    Close #N
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CargarFxs de TileEngine.bas")
End Sub

Sub CargarArrayLluvia()
On Error GoTo ErrHandler

    Dim N As Integer
    Dim I As Long
    Dim Nu As Integer

    N = FreeFile()
    Open App.path & "\init\fk.ind" For Binary Access Read As #N

    'cabecera
    Get #N, , MiCabecera

    'num de cabezas
    Get #N, , Nu

    'Resize array
    ReDim bLluvia(1 To Nu) As Byte

    For I = 1 To Nu
        Get #N, , bLluvia(I)
    Next I

    Close #N

  Exit Sub

ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CargarArrayLluvia de TileEngine.bas")
End Sub

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Byte, ByRef tY As Byte)

    tX = (UserPos.X - HalfWindowTileWidth) + viewPortX \ TilePixelWidth
    tY = (UserPos.Y - HalfWindowTileHeight) + viewPortY \ TilePixelHeight
    
    If tX <= 0 Then tX = 1
    If tY <= 0 Then tY = 1

End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
On Error GoTo ErrHandler
  
    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
        
    With charlist(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .active = 0 Then _
            NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)
        If Body < 0 Then
            Body = Not Body
            InitGrh .Body.Walk(1), Body
            InitGrh .Body.Walk(2), Body
            InitGrh .Body.Walk(3), Body
            InitGrh .Body.Walk(4), Body
        Else
            .Body = BodyData(Body)
        End If
        
        .Arma = WeaponAnimData(Arma)
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        
        .Heading = Heading
        
        'Reset moving stats
        If (CharIndex = UserCharIndex) Then
            .MoveOffsetX = g_Last_OffsetX
            .MoveOffsetY = g_Last_OffsetY
            g_Last_OffsetX = 0
            g_Last_OffsetY = 0
        Else
            .MoveOffsetX = 0
            .MoveOffsetY = 0
        End If
        
        .Moving = (.MoveOffsetX <> 0 Or .MoveOffsetY <> 0)
        
        If (.Moving) Then
            Call InitGrh(.Body.Walk(.Heading), .Body.Walk(.Heading).GrhIndex, 1)
            If (Not .UsandoArma) Then
                Call InitGrh(.Arma.WeaponWalk(.Heading), .Arma.WeaponWalk(.Heading).GrhIndex, 1)
            End If
            Call InitGrh(.Escudo.ShieldWalk(.Heading), .Escudo.ShieldWalk(.Heading).GrhIndex, 1)
        End If
        
        'Update position
        .Pos.X = X
        .Pos.Y = Y
        
        ' Create virtual sound source
        Set .SoundSource = Engine_Audio.CreateEmitter(X, Y)
        
        'Make active
        .active = 1
        .auraIndex = -1
        
        MapData(X, Y).CharIndex = CharIndex

        Call UpdateNodeSceneChar(CharIndex)
        Call Aurora_Scene.Insert(.Node)

      End With
      
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub MakeChar de TileEngine.bas")
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)
On Error GoTo ErrHandler
  
    Dim I As Byte
    
    With charlist(CharIndex)
        .active = 0
        .Alignment = eCharacterAlignment.Neutral
        .Criminal = 0
        .Atacable = False
        .auraIndex = -1
        For I = 0 To .LastFx
            .Fx(I).FxIndex = 0
        Next I
        
        .glowFxIndex = 0
        .invisible = False
#If EnableSecurity Then
        Call MI(CualMI).ResetInvisible(CharIndex)
#End If
        .Moving = 0
        .muerto = False
        .Nombre = ""
        .pie = False
        .Pos.X = 0
        .Pos.Y = 0
        .UsandoArma = False
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetCharInfo de TileEngine.bas")
End Sub


Sub EraseChar(ByVal CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
On Error GoTo ErrHandler
  
    charlist(CharIndex).active = 0
    
    'Update lastchar
    If CharIndex = LastChar Then
        Do Until charlist(LastChar).active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If

    Call Aurora_Scene.Remove(charlist(CharIndex).Node)
    
    MapData(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y).CharIndex = 0
    
    'Remove char's dialog
    Call Dialogos.RemoveDialog(CharIndex)
    
    If (CharIndex <> UserCharIndex) Then
        Call ResetCharInfo(CharIndex)
    End If
            
    ' Destroy virtual sound source
    Call Engine_Audio.DeleteEmitter(charlist(CharIndex).SoundSource, False)
        
    'Update NumChars
    NumChars = NumChars - 1
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EraseChar de TileEngine.bas")
End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2, Optional ByVal Manual As Boolean = True)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
On Error GoTo ErrHandler
  
    Grh.GrhIndex = GrhIndex
    If GrhIndex = 0 Then Exit Sub
    
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started And Manual Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.FrameTimer = timerEngine
    Grh.Speed = GrhData(Grh.GrhIndex).Speed
    Grh.Alpha = 255
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub InitGrh de TileEngine.bas")
End Sub

Public Sub InitGrhDepth(ByRef Grh As Grh, ByVal Layer As Long, ByVal X As Long, ByVal Y As Long, ByVal Z As Long)
    Grh.Depth = GetDepth(Layer, X, Y, Z)
End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
On Error GoTo ErrHandler
  
    Dim addX As Integer
    Dim addY As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.NORTH
                addY = -1
        
            Case E_Heading.EAST
                addX = 1
        
            Case E_Heading.SOUTH
                addY = 1
            
            Case E_Heading.WEST
                addX = -1
        End Select

        nX = X + addX
        nY = Y + addY

        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.Y = nY
                
        If (MapData(X, Y).CharIndex = CharIndex) Then
            MapData(X, Y).CharIndex = 0
        End If
          
        Call UpdateNodeSceneChar(CharIndex)
        Call Aurora_Scene.Update(.Node)

        .MoveOffsetX = -1 * (TilePixelWidth * addX)
        .MoveOffsetY = -1 * (TilePixelHeight * addY)

        Call Engine_Audio.UpdateEmitter(.SoundSource, nX, nY)

        .Moving = 1
        .Heading = nHeading
        
        Call InitGrh(.Body.Walk(.Heading), .Body.Walk(.Heading).GrhIndex, 1)
        If (Not .UsandoArma) Then
            Call InitGrh(.Arma.WeaponWalk(.Heading), .Arma.WeaponWalk(.Heading).GrhIndex, 1)
        End If
        Call InitGrh(.Escudo.ShieldWalk(.Heading), .Escudo.ShieldWalk(.Heading).GrhIndex, 1)

        .scrollDirectionX = addX
        .scrollDirectionY = addY
    End With
    
    If UserEstado = 0 Then Call DoPasosFx(CharIndex)

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub MoveCharbyHead de TileEngine.bas")
End Sub

Sub DoPasosFx(ByVal CharIndex As Integer)
    
    On Error GoTo ErrHandler

    With charlist(CharIndex)

        If Not .IsSailing Then

            If Not .muerto And (.priv = 0 Or .priv > 5) And Not charlist(UserCharIndex).muerto Then
                .pie = Not .pie
                
                Call Engine_Audio.PlayEffect(IIf(.pie, SND_PASOS1, SND_PASOS2), .SoundSource)
            End If

        Else
            ' TODO : Actually we would have to check if the CharIndex char is in the water or not....
            Call Engine_Audio.PlayEffect(SND_NAVEGANDO, .SoundSource)
        End If

    End With

    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoPasosFx de TileEngine.bas")

End Sub

Sub MoveCharByTelep(ByVal CharIndex As Integer, ByVal X As Integer, ByVal Y As Integer)
    With (charlist(CharIndex))
        If (MapData(.Pos.X, .Pos.Y).CharIndex = CharIndex) Then
            MapData(.Pos.X, .Pos.Y).CharIndex = 0
        End If
                    
        .Pos.X = X
        .Pos.Y = Y

        Call Engine_Audio.UpdateEmitter(.SoundSource, X, Y)
            
        MapData(X, Y).CharIndex = CharIndex

        Call UpdateNodeSceneChar(CharIndex)
        Call Aurora_Scene.Update(.Node)
        
    End With
End Sub

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
On Error GoTo ErrHandler
  
    Dim X As Integer
    Dim Y As Integer
    Dim addX As Integer
    Dim addY As Integer
    Dim nHeading As E_Heading
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        If (MapData(X, Y).CharIndex = CharIndex) Then
            MapData(X, Y).CharIndex = 0
        End If
        
        addX = nX - X
        addY = nY - Y
        
        If Sgn(addX) = 1 Then
            nHeading = E_Heading.EAST
        ElseIf Sgn(addX) = -1 Then
            nHeading = E_Heading.WEST
        ElseIf Sgn(addY) = -1 Then
            nHeading = E_Heading.NORTH
        ElseIf Sgn(addY) = 1 Then
            nHeading = E_Heading.SOUTH
        End If

        MapData(nX, nY).CharIndex = CharIndex
        
        .Pos.X = nX
        .Pos.Y = nY
    
        Call UpdateNodeSceneChar(CharIndex)
        Call Aurora_Scene.Update(.Node)

        .MoveOffsetX = -1 * (TilePixelWidth * addX)
        .MoveOffsetY = -1 * (TilePixelHeight * addY)

        Call Engine_Audio.UpdateEmitter(.SoundSource, nX, nY)
            
        .Moving = 1
        .Heading = nHeading
        
        Call InitGrh(.Body.Walk(.Heading), .Body.Walk(.Heading).GrhIndex, 1)
        If (Not .UsandoArma) Then
            Call InitGrh(.Arma.WeaponWalk(.Heading), .Arma.WeaponWalk(.Heading).GrhIndex, 1)
        End If
        Call InitGrh(.Escudo.ShieldWalk(.Heading), .Escudo.ShieldWalk(.Heading).GrhIndex, 1)
        
        .scrollDirectionX = Sgn(addX)
        .scrollDirectionY = Sgn(addY)
        
        'parche para que no medite cuando camina
        If .Fx(0).FxIndex = FxMeditar.CHICO Or .Fx(0).FxIndex = FxMeditar.GRANDE Or .Fx(0).FxIndex = FxMeditar.MEDIANO Or .Fx(0).FxIndex = FxMeditar.XGRANDE Or .Fx(0).FxIndex = FxMeditar.XXGRANDE Then
            .Fx(0).FxIndex = 0
        End If
    End With

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub MoveCharbyPos de TileEngine.bas")
End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
On Error GoTo ErrHandler
  
    Dim X As Integer
    Dim Y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            Y = -1
        
        Case E_Heading.EAST
            X = 1
        
        Case E_Heading.SOUTH
            Y = 1
        
        Case E_Heading.WEST
            X = -1
    End Select
    
    'Fill temp pos
    tX = UserPos.X + X
    tY = UserPos.Y + Y
    
    'Check to see if its out of bounds
    If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = X
        UserPos.X = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = True
        
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub MoveScreen de TileEngine.bas")
End Sub

''
' Loads grh data using the new file format.
'
' @return   True if the load was successfull, False otherwise.

Private Function LoadGrhData() As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim handle As Integer
    Dim fileVersion As Long
    
    'Open files
    handle = FreeFile()
    
    Open IniPath & "Graficos.ind" For Binary Access Read As handle
    Seek #1, 1
    
    'Get file version
    Get handle, , fileVersion
    
    'Get number of grhs
    Get handle, , grhCount
    
    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData
    
    While Not EOF(handle)
        Get handle, , Grh
   
        If Grh <> 0 Then
            With GrhData(Grh)
                'Get number of frames
                Get handle, , .NumFrames
                If .NumFrames <= 0 Then GoTo ErrorHandler
                
                ReDim .Frames(1 To GrhData(Grh).NumFrames)
                
                If .NumFrames > 1 Then
                    'Read a animation GRH set
                    For Frame = 1 To .NumFrames
                        Get handle, , .Frames(Frame)
                        If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                            GoTo ErrorHandler
                        End If
                    Next Frame
                    
                    Get handle, , .Speed
                    
                    If .Speed <= 0 Then GoTo ErrorHandler
                    
                    'Compute width and height
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    If .TileWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                    If .TileHeight <= 0 Then GoTo ErrorHandler
                Else
                    'Read in normal GRH data
                    Get handle, , .fileNum
                    If .fileNum <= 0 Then GoTo ErrorHandler
                    
                    Get handle, , GrhData(Grh).sX
                    If .sX < 0 Then GoTo ErrorHandler
                    
                    Get handle, , .sY
                    If .sY < 0 Then GoTo ErrorHandler
                    
                    Get handle, , .pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    Get handle, , .pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    Get handle, , .S0
                    Get handle, , .T0
                    Get handle, , .S1
                    Get handle, , .T1
                    
                    .S1 = .S0 + .S1
                    .T1 = .T0 + .T1
                
                    'Compute width and height
                    .TileWidth = .pixelWidth / TilePixelWidth
                    .TileHeight = .pixelHeight / TilePixelHeight
                    
                    .Frames(1) = Grh
                End If
            End With
        End If
    Wend
    
    Close handle

    LoadGrhData = True
Exit Function

ErrorHandler:
    Close handle
    LoadGrhData = False
End Function

Function MoveToLegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 01/08/2009
'Checks to see if a tile position is legal, including if there is a casper in the tile
'10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
'01/08/2009: ZaMa - Now invisible admins can't change position with caspers.
'*****************************************************************
On Error GoTo ErrHandler
  
    Dim CharIndex As Integer
    
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function
    End If
    
    CharIndex = MapData(X, Y).CharIndex
    '¿Hay un personaje?
    If CharIndex > 0 Then
    
        If MapData(UserPos.X, UserPos.Y).Blocked = 1 Then
            Exit Function
        End If
        
        With charlist(CharIndex)
            ' Si no es casper, no puede pasar
            If .iHead <> CASPER_HEAD And .iBody <> FRAGATA_FANTASMAL Then
                Exit Function
            Else
                ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
                If HayAgua(UserPos.X, UserPos.Y) Then
                    If Not HayAgua(X, Y) Then Exit Function
                Else
                    ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
                    If HayAgua(X, Y) Then Exit Function
                End If
                
                ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles
                If charlist(UserCharIndex).priv > 0 And charlist(UserCharIndex).priv < 6 Then
#If EnableSecurity Then
                    If MI(CualMI).IsInvisible(UserCharIndex) Then Exit Function
#Else
                    If charlist(UserCharIndex).invisible Then Exit Function
#End If
                End If
            End If
        End With
    End If
   
    If charlist(UserCharIndex).IsSailing <> HayAgua(X, Y) Then
        Exit Function
    End If
    
    MoveToLegalPos = True
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function MoveToLegalPos de TileEngine.bas")
End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean

    InMapBounds = (X >= XMinMapSize And X <= XMaxMapSize And Y >= YMinMapSize And Y <= YMaxMapSize)

End Function

Public Sub DrawGrh(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Z As Single, ByVal Center As Byte, ByVal Animate As Byte, Optional ByVal color As Long = -1, Optional ByVal killAtEnd As Byte = 1, Optional ByVal Angle As Integer = 0, Optional ByVal Alpha As Boolean = False)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Integer
    Dim CurrentFrame    As Integer

    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + ((timerEngine - Grh.FrameTimer) * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            Grh.FrameTimer = timerEngine

            
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                        If killAtEnd Then Exit Sub
                    End If
                End If
            End If
        End If
    End If
    
    If Grh.GrhIndex <= 0 Then
        Debug.Print X, Y
        Exit Sub
    End If
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
                    
        Dim Source As Math_Rectf, destination As Math_Rectf
        source.X1 = .S0
        source.Y1 = .T0
        source.X2 = .S1
        source.Y2 = .T1
        destination.X1 = X
        destination.Y1 = Y
        destination.X2 = X + .pixelWidth
        destination.Y2 = Y + .pixelHeight
        
        Call Draw(destination, source, Z, Angle, color, .fileNum, Alpha)
    End With

End Sub

Sub DrawGrhIndex(ByVal GrhIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Z As Single, ByVal Center As Byte, Optional ByVal color As Long = -1, Optional ByVal Angle As Integer = 0)
    If (GrhIndex = 0) Then Exit Sub
    
    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
                    
        Dim Source As Math_Rectf, destination As Math_Rectf
        source.X1 = .S0
        source.Y1 = .T0
        source.X2 = .S1
        source.Y2 = .T1
        destination.X1 = X
        destination.Y1 = Y
        destination.X2 = X + .pixelWidth
        destination.Y2 = Y + .pixelHeight
        
        Call Draw(destination, source, Z, Angle, color, .fileNum, False)
    End With
  
End Sub

Private Sub SetAlpha(ByRef Alpha As Single, ByVal IsActive As Boolean, ByVal Min As Byte, ByVal Max As Byte, ByVal Speed As Single)
On Error GoTo ErrHandler

        Dim value As Single
        
        If (IsActive) Then
            value = Alpha - Speed * timerElapsedTime
            If (value < Min) Then Alpha = Min Else Alpha = value
        Else
            value = Alpha + Speed * timerElapsedTime
            If (value > Max) Then Alpha = Max Else Alpha = value
        End If
        
        Exit Sub
        
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SetAlpha de TileEngine.bas")
End Sub

Sub RenderScreen(ByVal TileX As Integer, ByVal TileY As Integer, ByVal OffsetX As Integer, ByVal OffsetY As Integer)
    Dim ScreenMinY      As Long  'Start Y pos on current screen
    Dim ScreenMaxY      As Long  'End Y pos on current screen
    Dim ScreenMinX      As Long  'Start X pos on current screen
    Dim ScreenMaxX      As Long  'End X pos on current screen
    Dim MinY            As Long  'Start Y pos on current map
    Dim MaxY            As Long  'End Y pos on current map
    Dim MinX            As Long  'Start X pos on current map
    Dim MaxX            As Long  'End X pos on current map
    Dim X               As Long
    Dim Y               As Long
    Dim Drawable        As Long
    Dim DrawableX       As Long
    Dim DrawableY       As Long
    Dim DrawableType    As Long
    Dim IsAlphaActive   As Boolean

    'Figure out Ends and Starts of screen
    ScreenMinY = TileY - HalfWindowTileHeight
    ScreenMaxY = TileY + HalfWindowTileHeight
    ScreenMinX = TileX - HalfWindowTileWidth
    ScreenMaxX = TileX + HalfWindowTileWidth
    
    'Figure out Ends and Starts of map
    MinY = ScreenMinY
    MaxY = ScreenMaxY
    MinX = ScreenMinX
    MaxX = ScreenMaxX

    If OffsetY < 0 Then
        MaxY = MaxY + 1
    ElseIf OffsetY > 0 Then
        MinY = MinY - 1
    End If
    If OffsetX < 0 Then
        MaxX = MaxX + 1
    ElseIf OffsetX > 0 Then
        MinX = MinX - 1
    End If
    
    If MinY < YMinMapSize Then MinY = YMinMapSize
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    If MinX < XMinMapSize Then MinX = XMinMapSize
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize

    For Y = MinY To MaxY
        DrawableY = (Y - ScreenMinY) * TilePixelHeight + OffsetY
        
        For X = MinX To MaxX
            DrawableX = (X - ScreenMinX) * TilePixelWidth + OffsetX
        
            Call DrawGrh(MapData(X, Y).Graphic(1), DrawableX, DrawableY, GetDepth(1, X, Y), 0, 1)
        Next X
    Next Y
    
    Dim UserX As Long, UserY As Long
    UserX = charlist(UserCharIndex).Pos.X
    UserY = charlist(UserCharIndex).Pos.Y
    
    Dim Results() As Partitioner_Item
    
    ' Get the entities from the quadtree.
    Call Aurora_Scene.Query(MinX - 1, MinY - 1, MaxX + 1, MaxY + 1, Results)
    
    Call SetAlpha(RoofAlpha, bTecho, ROOF_ALPHA_MIN, ROOF_ALPHA_MAX, ROOF_ALPHA_SPEED)
    
    For Drawable = 0 To UBound(Results)
        With Results(Drawable)
            
            Y = .Y
            DrawableX = (.X - ScreenMinX) * TilePixelWidth + OffsetX
            DrawableY = (.Y - ScreenMinY) * TilePixelHeight + OffsetY
            DrawableType = .Type
            
            With MapData(.X, .Y)
                Select Case (DrawableType)
                    Case 2
                        Call DrawGrh(.Graphic(2), DrawableX, DrawableY, .Graphic(2).Depth, 1, 1)
                    Case 3
                        Call DrawGrh(.Graphic(3), DrawableX, DrawableY, .Graphic(3).Depth, 1, 1, , , , True)
                    Case 4
                        Call DrawGrh(.Graphic(4), DrawableX, DrawableY, .Graphic(4).Depth, 1, 1, RGBA(255, 255, 255, RoofAlpha), , , True)
                    Case 5
                        If .OBJInfo.CanBeTransparent Then
                            IsAlphaActive = (Y > UserY And Aurora_Scene.Overlaps(UserX, UserY, 1, Results(Drawable)))
                            Call SetAlpha(.ObjGrh.Alpha, IsAlphaActive, OBJECT_ALPHA_MIN, OBJECT_ALPHA_MAX, OBJECT_ALPHA_SPEED)
                        End If
                    
                        Call DrawGrh(.ObjGrh, DrawableX, DrawableY, .ObjGrh.Depth, 1, 1, RGBA(255, 255, 255, .ObjGrh.Alpha), , , True)
                    Case 6
                        Call CharRender(.CharIndex, DrawableX, DrawableY)
                End Select
            End With
        End With
   Next Drawable
    
   ' Call ExtendRender.Render(TileX, TileY, OffsetX, OffsetY)
End Sub

Public Function RenderSounds()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/30/2008
'Actualiza todos los sonidos del mapa.
'**************************************************************
On Error GoTo ErrHandler

    If bLluvia(PlayerData.CurrentMap.Number) = 1 Then
        If bRain Then
            If bTecho Then
                If frmMain.IsPlaying <> PlayLoop.plLluviain Then
                    If RainBufferIndex Then _
                        Call Engine_Audio.DisableSound(RainBufferIndex)
                    RainBufferIndex = Engine_Audio.PlayEffect("lluviain.wav", Nothing, True)
                    frmMain.IsPlaying = PlayLoop.plLluviain
                End If
            Else
                If frmMain.IsPlaying <> PlayLoop.plLluviaout Then
                    If RainBufferIndex Then _
                        Call Engine_Audio.DisableSound(RainBufferIndex)
                    RainBufferIndex = Engine_Audio.PlayEffect("lluviaout.wav", Nothing, True)
                    frmMain.IsPlaying = PlayLoop.plLluviaout
                End If
            End If
        End If
    End If

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function RenderSounds de TileEngine.bas")
End Function


Public Function InitTileEngine(ByVal setDisplayFormhWnd As Long, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal setTileHeight As Integer, ByVal setTileWidth As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer) As Boolean
'***************************************************
'Author: Aaron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Creates all DX objects and configures the engine to start running.
'***************************************************

    IniPath = App.path & "\Init\"
    
    'Fill startup variables

    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight


    HalfWindowTileHeight = setTileHeight \ 2
    HalfWindowTileWidth = setTileWidth \ 2

    'Set FPS value to 60 for startup
    FPS = 60
    FramesPerSecCounter = 60
    
    MinXBorder = XMinMapSize + HalfWindowTileWidth
    MaxXBorder = XMaxMapSize - HalfWindowTileWidth
    MinYBorder = YMinMapSize + HalfWindowTileHeight
    MaxYBorder = YMaxMapSize - HalfWindowTileHeight
    
    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set intial user position
    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY

    ' Initialize Aurora Engine
    Dim Aurora_Configuration As Aurora_Engine.Kernel_Properties
    Aurora_Configuration.WindowHandle = frmMain.picMain.hwnd
    Aurora_Configuration.WindowWidth = frmMain.picMain.ScaleWidth
    Aurora_Configuration.WindowHeight = frmMain.picMain.ScaleHeight
    Aurora_Configuration.WindowTitle = "Argentum Online"
    Aurora_Configuration.LogFilename = App.path & "/LOGS/Aurora.log"
    Call Kernel.Initialize(eKernelModeClient, Aurora_Configuration)

    Set CurrentUIDeviceCamera = New Graphic_Camera
    
    Set Aurora_Audio = Kernel.Audio
    Set Aurora_Content = Kernel.Content
    Set Aurora_Graphics = Kernel.Graphics
    Set Aurora_Network = Kernel.Network

    ' Initialize Content Manager
    Aurora_Content.AddLocator "Textures", New clsContentLocatorPacked
    Aurora_Content.AddSystemLocator "Resources", App.path & "/" 'TODO: Nightw: Change "/" for "/Resources/" and move the resources inside this folder
        
    ' Initialize Renderer
    Set Aurora_Renderer = Kernel.Renderer

    ' Initialize Techniques
    Set g_Technique_1 = Aurora_Content.Load("Resources://Resources/Pipeline/Sprite.effect", eResourceTypePipeline)
    Set g_Technique_2 = Aurora_Content.Load("Resources://Resources/Pipeline/Sprite_Alpha.effect", eResourceTypePipeline)

    ' Initialize Rain
    Set g_Rain_Material = Aurora_Content.Retrieve("Memory://Material/Base/Rain", eResourceTypeMaterial, True)
    Call g_Rain_Material.SetTexture(0, Aurora_Content.Load("Textures://15168.png", eResourceTypeTexture))
    Call g_Rain_Material.SetSampler(0, eTextureEdgeRepeat, eTextureEdgeRepeat, eTextureFilterTrilinear)
    Call Aurora_Content.Register(g_Rain_Material, False)

    RoofAlpha = ROOF_ALPHA_MAX
    
    Call LoadFontDescription

    Call LoadGrhData
    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
    Call CargarFxs
    Call LoadMasteries
    Set ExtendRender = New clsExtendRenderBase
    InitTileEngine = True
End Function

Public Sub DeinitTileEngine()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/14/07
'Destroys all DX objects
'***************************************************
On Error GoTo ErrHandler
  

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DeinitTileEngine de TileEngine.bas")
End Sub

Function ShowNextFrame(ByVal MouseViewX As Integer, ByVal MouseViewY As Integer) As Boolean

    '***************************************************
    'Author: Arron Perkins
    'Last Modification: 08/14/07
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Updates the game's model and renders everything.
    '***************************************************

    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
                
    Static FrameTime      As Currency
    Static FrameNextTime  As Currency

    If EngineRun Then
    
        FrameTime = FrameTime + GetFrameElapsedTime()
        
        If (Not GameConfig.Graphics.bUseVerticalSync Or FrameTime >= FrameNextTime) Then
            
            If UserMoving Then
    
                '****** Move screen Left and Right if needed ******
                If AddtoUserPos.X <> 0 Then
                    OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame
    
                    If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                        OffsetCounterX = 0
                        AddtoUserPos.X = 0
                        UserMoving = False
                    End If
    
                End If
                
                '****** Move screen Up and Down if needed ******
                If AddtoUserPos.Y <> 0 Then
                    OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame
    
                    If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                        OffsetCounterY = 0
                        AddtoUserPos.Y = 0
                        UserMoving = False
                    End If
    
                End If
    
            End If
            
            'Update mouse position within view area
            Call ConvertCPtoTP(MouseViewX, MouseViewY, MouseTileX, MouseTileY)
    
            '****** Update screen ******
            Dim Viewport As Math_Rectf
            Viewport.X1 = 0: Viewport.X2 = frmMain.picMain.ScaleWidth
            Viewport.Y1 = 0: Viewport.Y2 = frmMain.picMain.ScaleHeight
            Call Aurora_Graphics.Prepare(&H0, Viewport, eClearAll, 0, 1#, 0)

            Dim Camera As New Graphic_Camera
            Call Camera.SetOrthographic(0, frmMain.picMain.ScaleWidth, frmMain.picMain.ScaleHeight, 0, 1000, -1000)
            Call Camera.Compute ' TODO: Camera movement
            
            Call Aurora_Renderer.Begin(Camera, timerEngine / 1000#)

            Call SetEffect(UserEstado = 1)

            If (Not UserCiego) Then
                Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
            End If
        
            Call Dialogos.Render
            Call DialogosClanes.Draw(FuentesJuego.Guild)
            Call DibujarCartel
                        
            If (bRain And bLluvia(PlayerData.CurrentMap.Number)) Then
                Call DrawRain
            End If
                        
            If GameConfig.Extras.bShowFPS Then
                        
                DrawText frmMain.picMain.ScaleWidth - 4, 0, 0#, "FPS: " & FramesPerSecCounter, &HFF0000FF, eRendererAlignmentRightBaseline, _
                        FuentesJuego.Info
            End If
            
            If GameConfig.Extras.bShowLatency Then
                DrawText frmMain.picMain.ScaleWidth - 4, FuentesJuego.Info.Tamanio + 4, 0#, "PING: " & currentPingTime, &HFFFFFFFF, eRendererAlignmentRightBaseline, _
                            FuentesJuego.Info
            End If
                            
            ' Draw quest related requirements on screen
            If PlayerData.Guild.Quest.Id > 0 Then
                Call DrawQuestStageRequirements
            End If

            Call Aurora_Renderer.End
            Call Aurora_Graphics.Commit(&H0, False)

            FrameNextTime = FrameTime + (1000 / 144) ' 144 = Target FPS
            
            ShowNextFrame = True
                
            'Get timing info
            timerElapsedTime = GetElapsedTime()
            timerTicksPerFrame = timerElapsedTime * ENGINE_SPEED
            timerEngine = timerEngine + timerElapsedTime
            
        End If

    End If

End Function

Public Sub DrawQuestStageRequirements()
    Dim CurrentHeight As Integer
    Dim PositionStep As Integer
    Dim I As Integer
    
    PositionStep = FuentesJuego.Info.Tamanio
    CurrentHeight = 0
    
    With PlayerData.Guild.Quest.CurrentStageProgress
        For I = 1 To .PreCalculatedScreenTextLineQty
            DrawText 5, CurrentHeight + PositionStep, 0#, .PreCalculatedScreenText(I).Text, .PreCalculatedScreenText(I).Color, eRendererAlignmentLeftBaseline, FuentesJuego.Info, True
            CurrentHeight = CurrentHeight + PositionStep
        Next I
    End With
    
End Sub

Public Sub DrawText(ByVal X As Long, ByVal Y As Long, ByVal Z As Single, ByRef Word As String, ByVal Color As Long, ByVal Alignment As Renderer_Alignment, ByRef Font As tFuente, Optional ByVal Outline As Boolean = 0)
    Call Aurora_Renderer.DrawFont(Font.Asset, Word, X, Y, Z, Font.Tamanio, Color, Alignment)
End Sub

Private Function GetFrameElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
On Error GoTo ErrHandler
  
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        Call QueryPerformanceFrequency(timer_freq)
    End If
    If end_time = 0 Then
        Call QueryPerformanceCounter(end_time)
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetFrameElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetFrameElapsedTime de TileEngine.bas")
End Function

Private Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
On Error GoTo ErrHandler
  
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        Call QueryPerformanceFrequency(timer_freq)
    End If
    If end_time = 0 Then
        Call QueryPerformanceCounter(end_time)
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetElapsedTime de TileEngine.bas")
End Function

Public Function GetInputElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
On Error GoTo ErrHandler
  
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        Call QueryPerformanceFrequency(timer_freq)
    End If
    If end_time = 0 Then
        Call QueryPerformanceCounter(end_time)
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetInputElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetElapsedTime de TileEngine.bas")
End Function

Private Sub CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Single, ByVal PixelOffsetY As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 25/05/2011 (Amraphen)
'Draw char's to screen without offcentering them
'16/09/2010: ZaMa - Ya no se dibujan los bodies cuando estan invisibles.
'25/05/2011: Amraphen - Agregado movimiento de armas al golpear.
'***************************************************
On Error GoTo ErrHandler
  
    Dim moved As Boolean
    Dim attacked As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim color As Long
    Dim I As Byte
    Dim LastIndex As Byte
    Dim TextOffsetY As Integer
    Dim Alpha As Long
    Dim OverheadIcon As Integer
    
    Alpha = &H60FFFFFF
    
    With charlist(CharIndex)
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame

                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
            
            Call Engine_Audio.UpdateEmitter(.SoundSource, .Pos.X, .Pos.Y)
        End If

        If .UsandoArma And .Arma.WeaponWalk(.Heading).Started Then _
            attacked = True
            
        'If done moving stop animation
        If Not moved Then
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
            .Body.Walk(.Heading).FrameTimer = timerEngine
            
            If Not attacked Then
                .Arma.WeaponWalk(.Heading).Started = 0
                .Arma.WeaponWalk(.Heading).FrameCounter = 1
                .Arma.WeaponWalk(.Heading).FrameTimer = timerEngine
            
                .UsandoArma = False
            End If
            
            .Escudo.ShieldWalk(.Heading).Started = 0
            .Escudo.ShieldWalk(.Heading).FrameCounter = 1
            .Escudo.ShieldWalk(.Heading).FrameTimer = timerEngine
                
            .Moving = False
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        Dim fuente As tFuente
        ' Set a default font.
        fuente = FuentesJuego.FuenteBase
        
#If EnableSecurity Then
        If Not MI(CualMI).IsInvisible(CharIndex) Then
#Else
        If Not .invisible Then
#End If
            'Draw Body
            If .Body.Walk(.Heading).GrhIndex Then _
                Call DrawGrh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, GetDepth(3, .Pos.X, .Pos.Y, 2), 1, 1, , 0, , True)
            
            'Draw Head
            If .Head.Head(.Heading).GrhIndex Or .NpcNumber > 0 Then
                If .Head.Head(.Heading).GrhIndex Then
                    Call DrawGrh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, GetDepth(3, .Pos.X, .Pos.Y, 3), 1, 0, , , , True)
                End If
                
                'Draw Helmet
                If .Casco.Head(.Heading).GrhIndex Then
                    Call DrawGrh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, GetDepth(3, .Pos.X, .Pos.Y, 4), 1, 0, , , , True)
                End If
                
                'Draw Weapon
                If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                    Call DrawGrh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, GetDepth(3, .Pos.X, .Pos.Y, 6), 1, 1, , 0, , True)
                
                'Draw Shield
                If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                    Call DrawGrh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, GetDepth(3, .Pos.X, .Pos.Y, 5), 1, 1, , 0, , True)
                

                 
                'Draw name over head
                If LenB(.Nombre) > 0 Then
                    If (Nombres = eNombresView.Rollover _
                        And (Abs(MouseTileX - .Pos.X) < 2 And (Abs(MouseTileY - .Pos.Y)) < 2)) _
                        Or (Nombres = eNombresView.Fixed And (esGM(UserCharIndex) Or _
                        (.NpcNumber = 0 Or Abs(MouseTileX - .Pos.X) < 2 And (Abs(MouseTileY - .Pos.Y)) < 2))) Then
                        Pos = getTagPosition(.Nombre)
                        
                        If .NpcNumber = 0 Then
                            If .priv = 0 Then
                                ' Is a normal user
                                fuente = GetFontByAlignment(.Alignment)
                            Else
                                ' Is a an admin
                                fuente = GetFontByPrivs(.priv)
                            End If
                        Else
                            fuente = FuentesJuego.NickNpcs
                        End If
                        
                        line = Left$(.Nombre, Pos - 2)
                        Call DrawText(PixelOffsetX + TilePixelWidth \ 2, PixelOffsetY + 35, GetDepth(3, .Pos.X, .Pos.Y, 8), line, fuente.Color, eRendererAlignmentCenterTop, fuente, True)
    
                        'Clan
                        line = mid$(.Nombre, Pos)
                        Call DrawText(PixelOffsetX + TilePixelWidth \ 2, PixelOffsetY + 50, GetDepth(3, .Pos.X, .Pos.Y, 8), line, fuente.Color, eRendererAlignmentCenterTop, fuente, True)
                    End If
                End If
            End If
        Else
            If .UseInvisibilityAlpha Or CharIndex = UserCharIndex Or esGM(UserCharIndex) Then
            
                        'Draw Body
                    If .Body.Walk(.Heading).GrhIndex Then _
                        Call DrawGrh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, GetDepth(3, .Pos.X, .Pos.Y, 2), 1, 1, Alpha, 0, , True)
            
                    'Draw Head
                    If .Head.Head(.Heading).GrhIndex Or .NpcNumber > 0 Then
                        If .Head.Head(.Heading).GrhIndex Then
                            Call DrawGrh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, GetDepth(3, .Pos.X, .Pos.Y, 3), 1, 0, Alpha, , , True)
                        End If
                
                        'Draw Helmet
                        If .Casco.Head(.Heading).GrhIndex Then
                            Call DrawGrh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, GetDepth(3, .Pos.X, .Pos.Y, 4), 1, 0, Alpha, , , True)
                        End If
                
                        'Draw Weapon
                        If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                            Call DrawGrh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, GetDepth(3, .Pos.X, .Pos.Y, 6), 1, 1, Alpha, 0, , True)
                
                        'Draw Shield
                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                            Call DrawGrh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, GetDepth(3, .Pos.X, .Pos.Y, 5), 1, 1, Alpha, 0, , True)
                
                
                        'Draw name over head
                        If LenB(.Nombre) > 0 Then
                            If (Nombres = eNombresView.Rollover _
                                And (Abs(MouseTileX - .Pos.X) < 2 And (Abs(MouseTileY - .Pos.Y)) < 2)) _
                                Or (Nombres = eNombresView.Fixed And (.NpcNumber = 0 Or Abs(MouseTileX - .Pos.X) < 2 _
                                And (Abs(MouseTileY - .Pos.Y)) < 2)) Then
                                Pos = getTagPosition(.Nombre)
                        
                                If .NpcNumber = 0 Then
                                    If .priv = 0 Then
                                        ' Is a normal user
                                        fuente = GetFontByAlignment(.Alignment)
                                    Else
                                        ' Is a an admin
                                        fuente = GetFontByPrivs(.priv)
                                    End If
                                Else
                                    fuente = FuentesJuego.NickNpcs
                                End If
                        
                                line = Left$(.Nombre, Pos - 2)
                                Call DrawText(PixelOffsetX + TilePixelWidth \ 2, PixelOffsetY + 35, GetDepth(3, .Pos.X, .Pos.Y, 8), line, fuente.Color, eRendererAlignmentCenterTop, fuente, True)
    
                                'Clan
                                line = mid$(.Nombre, Pos)
                                Call DrawText(PixelOffsetX + TilePixelWidth \ 2, PixelOffsetY + 50, GetDepth(3, .Pos.X, .Pos.Y, 8), line, fuente.Color, eRendererAlignmentCenterTop, fuente, True)
                            End If
                        End If
                    End If
            End If
        End If
        
        ' Set chat text offsets
        TextOffsetY = GetChatOverheadTextOffset(CharIndex, PixelOffsetY, TilePixelHeight)
        
        'Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffsetX + TilePixelWidth \ 2, TextOffsetY, 0#, CharIndex)
        'If Not (.overHeadAnimation Is Nothing) Then
            'TODO use global var for tickcounts
       '     .overHeadAnimation.Tick GetTickCount()
       '     .overHeadAnimation.Draw PixelOffsetX, PixelOffsetY
       '     If .overHeadAnimation.IsComplete Then
       '         Set .overHeadAnimation = Nothing
       '     End If
       ' End If
        'Draw FX

        For I = 0 To .LastFx
            If .Fx(I).FxIndex <> 0 Then
                LastIndex = I
                Call DrawGrh(.Fx(I).FxGrh, PixelOffsetX + FxData(.Fx(I).FxIndex).OffsetX, PixelOffsetY + FxData(.Fx(I).FxIndex).OffsetY, GetDepth(3, .Pos.X, .Pos.Y, 7), 1, 1, , , , True)
                
                'Check if animation is over
                If .Fx(I).FxGrh.Started = 0 Then _
                    .Fx(I).FxIndex = 0
            End If
        Next I
        
        If .LastFx > LastIndex Then
            .LastFx = LastIndex
        End If
                'Draw FX
        If .glowFxIndex <> 0 Then
            Call DrawGrh(.glowfX, PixelOffsetX + FxData(.glowFxIndex).OffsetX, PixelOffsetY + FxData(.glowFxIndex).OffsetY, GetDepth(3, .Pos.X, .Pos.Y, 7), 1, 1, , , , True)
            
            'Check if animation is over
            If .glowfX.Started = 0 Then _
                .glowFxIndex = 0
        End If
        
        If modQuests_Guild.IsQuestNpc(.NpcNumber) And PlayerData.Guild.Quest.CurrentStageProgress.RequirementsCompleted Then
            OverheadIcon = QUEST_OVERHEADICON
        Else
            OverheadIcon = .OverheadIcon
        End If
        
        ' Draw the overhead icon on top of the character (both player and npc) to determine the type of action that can be performed.
        If OverheadIcon <> 0 Then
            Dim GrhIcon As Grh
            GrhIcon.GrhIndex = OverheadIcon
            GrhIcon.FrameCounter = 1
            
            Call DrawGrh(GrhIcon, PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD - (OFFSET_HEAD / -2), GetDepth(3, .Pos.X, .Pos.Y, 4), 1, 0, , , , True)
        End If
        
    End With
  
    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CharRender de TileEngine.bas")
End Sub

Public Function GetChatOverheadTextOffset(ByVal CharIndex As Integer, ByVal PixelOffsetY As Integer, ByVal TilePixelHeight As Integer) As Integer

    Dim HeadGrhIndex As Long
    Dim BodyGrhIndex As Long

    With charlist(CharIndex)
        HeadGrhIndex = .Head.Head(.Heading).GrhIndex

        If (HeadGrhIndex > 0) Then
            ' Set the offset for the chat overhead
            GetChatOverheadTextOffset = (PixelOffsetY + .Body.HeadOffset.Y) - GrhData(HeadGrhIndex).pixelHeight + (TilePixelHeight \ 2)
        Else
            BodyGrhIndex = .Body.Walk(.Heading).GrhIndex
            
            If BodyGrhIndex > 0 Then
                ' If the user has no head, then we set the text offset chat overhead to be based on the body size.
                GetChatOverheadTextOffset = (PixelOffsetY + .Body.HeadOffset.Y) '- Int(GrhData(BodyGrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
                'GetChatOverheadTextOffset = (PixelOffsetY) - Int(GrhData(BodyGrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight - .Body.HeadOffset.Y
            End If
        End If

    End With

End Function


Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal Fx As Integer, ByVal Loops As Integer, Optional ByVal Slot As Byte = 255)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'11/10/2016: Anagrama - It now has layers.
'***************************************************
On Error GoTo ErrHandler
  
    Dim I As Byte
    Dim MyIndex As Byte
    Dim cantFx As Byte
    
    With charlist(CharIndex)
        If Slot = 255 Then
            If Fx = 0 Then Exit Sub
            
            For I = 1 To .LastFx
                If .Fx(I).FxIndex = 0 Then
                    MyIndex = I
                    Exit For
                End If
            Next I
            
            If MyIndex = 0 Then
                If .LastFx < 10 Then
                    .LastFx = .LastFx + 1
                    MyIndex = .LastFx
                Else
                    MyIndex = 10
                End If
            End If
        Else
            cantFx = UBound(.Fx)
            If Slot = 0 And Fx = 0 Then
                'reset all fx
                For I = 0 To cantFx
                        .Fx(I).FxGrh.FrameTimer = 0
                        .Fx(I).FxGrh.FrameCounter = 0
                        .Fx(I).FxGrh.GrhIndex = 0
                        .Fx(I).FxGrh.Loops = 0
                        .Fx(I).FxGrh.Speed = 0
                        .Fx(I).FxGrh.Started = 0
                        .Fx(I).FxIndex = 0
                Next I
                .LastFx = 0
            End If
            'check for ilegal access
            If Slot <= cantFx Then
                MyIndex = Slot
            Else
                Exit Sub
            End If
        End If
        
        .Fx(MyIndex).FxIndex = Fx
        
        If .Fx(MyIndex).FxIndex > 0 Then
            Call InitGrh(.Fx(MyIndex).FxGrh, FxData(Fx).Animacion)
        
            .Fx(MyIndex).FxGrh.Loops = Loops
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SetCharacterFx de TileEngine.bas")
End Sub

Public Sub SetCharacterGlowFx(ByVal CharIndex As Integer, ByVal Fx As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************
On Error GoTo ErrHandler
  
    With charlist(CharIndex)
        .glowFxIndex = Fx
        
        If .glowFxIndex > 0 Then
            Call InitGrh(.glowfX, FxData(Fx).Animacion)
        
            .glowfX.Loops = Loops
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SetCharacterGlowFx de TileEngine.bas")
End Sub


Public Function GetFontByAlignment(ByVal Alignment As eCharacterAlignment) As tFuente
On Error GoTo ErrHandler
  
    Select Case Alignment
        Case eCharacterAlignment.Newbie
            GetFontByAlignment = FuentesJuego.NickNewbie
        Case eCharacterAlignment.Neutral
            GetFontByAlignment = FuentesJuego.NickNeutral
        Case eCharacterAlignment.FactionLegion
            GetFontByAlignment = FuentesJuego.NickFactionLegion
        Case eCharacterAlignment.FactionRoyal
            GetFontByAlignment = FuentesJuego.NickFactionRoyal
        Case Else
            GetFontByAlignment = FuentesJuego.NickNeutral
    End Select
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetFontByPrivs de ModDDEX.bas")
End Function

Public Function GetFontByPrivs(ByVal privs As Integer) As tFuente
On Error GoTo ErrHandler
  
    Select Case privs
        Case 1
            GetFontByPrivs = FuentesJuego.NickConsejero
        Case 2
            GetFontByPrivs = FuentesJuego.NickSemidios
        Case 3
            GetFontByPrivs = FuentesJuego.NickDios
        Case 6
            GetFontByPrivs = FuentesJuego.NickConcilio
        Case 7
            GetFontByPrivs = FuentesJuego.NickConsejo
    End Select
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetFontByPrivs de ModDDEX.bas")
End Function

Public Function CreateFont(ByRef Asset As Graphic_Font, ByVal Tamanio As Long, ByVal Color As Long) As tFuente
    
    Set CreateFont.Asset = Asset
    CreateFont.Tamanio = Tamanio
    CreateFont.color = color
    
End Function

Public Sub LoadFontDescription()
On Error GoTo ErrHandler
    
    Dim Font As Graphic_Font
    Set Font = Aurora_Content.Load("Resources/Font/Primary.arfont", eResourceTypeFont)

    ' RGBA TODO: NIGHTW -> RGBA(R, G, B, A)
    FuentesJuego.FuenteBase = CreateFont(Font, 13, RGBA(255, 255, 255, 255))
    
    FuentesJuego.NickNewbie = CreateFont(Font, 12, RGBA(242, 192, 41, 255))
    FuentesJuego.NickNeutral = CreateFont(Font, 12, RGBA(184, 182, 176, 255))
    FuentesJuego.NickFactionRoyal = CreateFont(Font, 12, RGBA(0, 128, 255, 255))
    FuentesJuego.NickFactionLegion = CreateFont(Font, 12, RGBA(255, 0, 0, 255))
    
    FuentesJuego.NickAdmins = CreateFont(Font, 12, RGBA(255, 255, 255, 255))
    FuentesJuego.NickDios = CreateFont(Font, 12, RGBA(250, 250, 150, 255))
    FuentesJuego.NickSemidios = CreateFont(Font, 12, RGBA(30, 255, 48, 255))
    FuentesJuego.NickConsejero = CreateFont(Font, 12, RGBA(30, 150, 48, 255))
    FuentesJuego.NickAdmins = CreateFont(Font, 12, RGBA(180, 180, 180, 255))
    FuentesJuego.NickConcilio = CreateFont(Font, 12, RGBA(255, 50, 0, 255))
    FuentesJuego.NickConsejo = CreateFont(Font, 12, RGBA(240, 195, 255, 255))
    FuentesJuego.NickNpcs = CreateFont(Font, 12, RGBA(182, 169, 81, 255))
    
    FuentesJuego.Talk = CreateFont(Font, 13, RGBA(255, 255, 255, 255))
    FuentesJuego.Fight = CreateFont(Font, 13, RGBA(255, 0, 0, 255))
    FuentesJuego.Warning = CreateFont(Font, 13, RGBA(32, 51, 233, 255))
    FuentesJuego.Info = CreateFont(Font, 13, RGBA(65, 190, 156, 255))
    FuentesJuego.InfoBold = CreateFont(Font, 13, RGBA(49, 190, 156, 255))
    FuentesJuego.Execution = CreateFont(Font, 13, RGBA(130, 130, 130, 255))
    FuentesJuego.Party = CreateFont(Font, 13, RGBA(255, 180, 255, 255))
    FuentesJuego.Poison = CreateFont(Font, 3, RGBA(0, 255, 0, 255))
    
    FuentesJuego.Guild = CreateFont(Font, 13, RGBA(255, 255, 255, 255))
    FuentesJuego.Server = CreateFont(Font, 13, RGBA(0, 185, 0, 255))
    FuentesJuego.GuildMsg = CreateFont(Font, 13, RGBA(255, 199, 27, 255))
    
    FuentesJuego.ConsejoVesA = CreateFont(Font, 13, RGBA(0, 200, 255, 255))
    FuentesJuego.ConcilioVesA = CreateFont(Font, 13, RGBA(255, 50, 0, 255))

    FuentesJuego.GMSG = CreateFont(Font, 13, RGBA(255, 255, 255, 255))

    FuentesJuego.Inventarios = CreateFont(Font, 10, RGBA(255, 255, 255, 255))
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadFontDescription de Mod_TileEngine.bas")
End Sub

Public Sub DrawRain()
    Dim Animation As Single
    Animation = timerEngine / 1000#
    
    Dim Source As Math_Rectf
    Source.X1 = 0: Source.X2 = 1#
    Source.Y1 = 1 + Animation: Source.Y2 = Animation
    
    Dim destination As Math_Rectf
    destination.X1 = 0: destination.X2 = frmMain.picMain.ScaleWidth
    destination.Y1 = 0: destination.Y2 = frmMain.picMain.ScaleHeight
    
    Call Aurora_Renderer.DrawTexture(destination, Source, 0#, 0#, eRendererOrderOpaque, -1, g_Technique_1, g_Rain_Material)
End Sub

Public Sub Draw(ByRef destination As Math_Rectf, ByRef Source As Math_Rectf, ByVal Depth As Single, ByVal Angle As Single, ByVal Color As Long, ByVal Graphic As Long, ByVal Alpha As Boolean)
    Dim Material As Graphic_Material
    Set Material = Aurora_Content.Retrieve("Memory://Material://Base/" + CStr(Graphic), eResourceTypeMaterial, True)
    
    ' Create the Material on Demand
    If (Material.GetStatus = eResourceStatusNone) Then
        Dim Texture As Graphic_Texture
        Set Texture = Aurora_Content.Load("Textures://" + CStr(Graphic) + ".png", eResourceTypeTexture)
        
        If (Texture.GetStatus <> eResourceStatusLoaded) Then
            Debug.Print "Tile_Engine::Draw", "Failed to acquire texture"
            Exit Sub
        End If

        Call Material.SetTexture(0, Texture)
        
        Call Aurora_Content.Register(Material, False)
    End If

    If (Alpha) Then
        Call Aurora_Renderer.DrawTexture(destination, Source, Depth, Angle, eRendererOrderNormal, Color, g_Technique_2, Material)
    Else
        Call Aurora_Renderer.DrawTexture(destination, Source, Depth, Angle, eRendererOrderOpaque, Color, g_Technique_1, Material)
    End If
    
End Sub

Public Function GetCharacterDimension(ByVal CharIndex As Integer, ByRef RangeX As Single, ByRef RangeY As Single)
    Dim I As Long
    
    Dim BestX As Long
    Dim BestY As Long
            
    With charlist(CharIndex)
    
        ' Try to calculate the best width and height using all four direction of the entity's body
        If (.iBody <> 0) Then
            For I = 1 To 4
                If (GrhData(.Body.Walk(I).GrhIndex).TileWidth > RangeX) Then
                    RangeX = GrhData(.Body.Walk(I).GrhIndex).TileWidth
                End If
                If (GrhData(.Body.Walk(I).GrhIndex).TileHeight > RangeY) Then
                    RangeY = GrhData(.Body.Walk(I).GrhIndex).TileHeight
                End If
            Next I
        End If
                
        ' Try to calculate the best width and height using all four direction of the entity's body
        If (.iHead <> 0) Then

            For I = 1 To 4
                If (GrhData(.Head.Head(I).GrhIndex).TileWidth > RangeX) Then
                    RangeX = GrhData(.Head.Head(I).GrhIndex).TileWidth
                End If
                If (GrhData(.Head.Head(I).GrhIndex).TileHeight > BestY) Then
                    BestY = GrhData(.Head.Head(I).GrhIndex).TileHeight
                End If
            Next I

            RangeY = RangeY + BestY
        End If
            
        If (.Nombre <> vbNullString) Then
            RangeY = RangeY + 2
            
            BestX = Len(GetRawName(.Nombre)) * FuentesJuego.NickNpcs.Tamanio / 32
            If (BestX > RangeX) Then RangeX = BestX
        End If
        
            
        ' FX Too!
        BestX = 0
        BestY = 0
        
        For I = 0 To 9
            If (.Fx(I).FxIndex <> 0) Then
                If (GrhData(.Fx(I).FxGrh.GrhIndex).TileWidth > BestX) Then
                    BestX = GrhData(.Fx(I).FxGrh.GrhIndex).TileWidth
                End If
                If (GrhData(.Fx(I).FxGrh.GrhIndex).TileHeight > BestY) Then
                    BestY = GrhData(.Fx(I).FxGrh.GrhIndex).TileHeight
                End If
            End If
        Next I
        
        If (RangeX < BestX) Then RangeX = BestX
        If (RangeY < BestY) Then RangeY = BestY

    End With

End Function

Public Sub UIBegin(ByVal Device As Long, ByVal Width As Long, ByVal Height As Long, ByVal Tint As Long)
    CurrentUIDevice = Device
    
    Dim Viewport As Math_Rectf
    Viewport.X1 = 0: Viewport.X2 = Width
    Viewport.Y1 = 0: Viewport.Y2 = Height
    
    Call Aurora_Graphics.Prepare(Device, Viewport, eClearAll, Tint, 1#, 0)
    
    Call CurrentUIDeviceCamera.SetOrthographic(0, Width, Height, 0, 1, -1)
    Call CurrentUIDeviceCamera.Compute
    
    Call Aurora_Renderer.Begin(CurrentUIDeviceCamera, timerEngine / 1000#)
End Sub

Public Sub UIEnd()
    Call Aurora_Renderer.End
    
    Call Aurora_Graphics.Commit(CurrentUIDevice, False)
End Sub

Public Function GetDepth(ByVal Layer As Single, Optional ByVal X As Single = 1, Optional ByVal Y As Single = 1, Optional ByVal Z As Single = 1) As Single

    GetDepth = -1# + (Layer * 0.1) + ((Y - 1) * 0.001) + ((X - 1) * 0.00001) + ((Z - 1) * 0.000001)
    
End Function

Public Function RGBA(ByVal red As Long, ByVal green As Long, ByVal blue As Long, ByVal Alpha As Long) As Long
    If Alpha > 127 Then
        RGBA = RGB(red, green, blue) Or (Alpha - 128) * &H1000000 Or &H80000000
    Else
        RGBA = RGB(red, green, blue) Or Alpha * &H1000000
    End If
End Function

Public Sub SetEffect(ByVal Grayscale As Boolean)
    Dim Effect As Long
    
    If (Grayscale) Then
        Effect = 1
    Else
        Effect = 0
    End If
    
    Call Aurora_Renderer.SetParameter(0, Effect, 0, 0, 0)
End Sub

Public Function CreateMaterialWithTextureFromFile(ByRef FilePath As String) As Integer
    ' TODO: Wolfteni
    'CreateMaterialWithTextureFromFile = wGL_Graphic_Renderer.Create_Material
    
    'Call wGL_Graphic_Renderer.Update_Material_Texture(CreateMaterialWithTextureFromFile, &H0, wGL_Graphic.Create_Texture_From_Image(LoadBytesAbsolutePath(filePath)))

End Function

Public Sub DrawMaterial(ByRef destination As Math_Rectf, ByRef Source As Math_Rectf, ByVal Depth As Single, ByVal Angle As Single, ByVal Color As Long, ByVal Material As Integer, Optional ByVal Alpha As Boolean = False)
    ' TODO: Wolftein
    'If (Alpha) Then
    '    Call wGL_Graphic_Renderer.Draw(Destination, Source, Depth, Angle, Color, Material, g_Technique_2)
    'Else
    '    Call wGL_Graphic_Renderer.Draw(Destination, Source, Depth, Angle, Color, Material, g_Technique_1)
    'End If
    
End Sub
