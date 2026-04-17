VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExtendRenderBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MaxOverGrh As Long = 200
Private Const FadeStepTime As Long = 10

Private Type OverGrh
    Grh As Grh
    X As Long
    Y As Long
    Alpha As Double
    AlphaFactor As Double 'Used to decrease alpha
    Z As Double
End Type

Private OverGrhList(MaxOverGrh) As OverGrh
Private LastTick As Long
Private Enable As Boolean
Public Sub Render(ByVal TileX As Integer, ByVal TileY As Integer, ByVal OffsetX As Integer, ByVal OffsetY As Integer)
    If Enable = False Then
        Exit Sub
    End If
    Dim ScreenMinY  As Long  'Start Y pos on current screen
    Dim ScreenMaxY  As Long  'End Y pos on current screen
    Dim ScreenMinX  As Long  'Start X pos on current screen
    Dim ScreenMaxX  As Long  'End X pos on current screen
    Dim MinY        As Long  'Start Y pos on current map
    Dim MaxY        As Long  'End Y pos on current map
    Dim MinX        As Long  'Start X pos on current map
    Dim MaxX        As Long  'End X pos on current map
    Dim X           As Long
    Dim Y           As Long
    Dim Drawable    As Long
    Dim DrawableX   As Long
    Dim DrawableY   As Long



    'Figure out Ends and Starts of screen
    ScreenMinY = TileY - Mod_TileEngine.HalfWindowTileHeight
    ScreenMaxY = TileY + Mod_TileEngine.HalfWindowTileHeight
    ScreenMinX = TileX - Mod_TileEngine.HalfWindowTileWidth
    ScreenMaxX = TileX + Mod_TileEngine.HalfWindowTileWidth

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



    Dim ticks As Long
    Dim AlphaSwitch As Double

    ticks = GetTickCount()
    If ticks - LastTick >= FadeStepTime Then
        LastTick = ticks
        AlphaSwitch = 1
    Else
        AlphaSwitch = 0
    End If

    Dim OverGrhTotalDisable As Long
    OverGrhTotalDisable = 0

    Dim OverGrhIndex As Long
    For OverGrhIndex = 0 To MaxOverGrh - 1
        With OverGrhList(OverGrhIndex)
            If .Alpha > 0 Then
                X = .X
                Y = .Y

                If X >= MinX And X <= MaxX And Y >= MinY And Y <= MaxY Then
                    Dim AlphaColor As Long
                    AlphaColor = RGBA(255, 255, 255, CInt(.Alpha))
                    DrawableY = (Y - ScreenMinY) * Mod_TileEngine.TilePixelHeight + OffsetY
                    DrawableX = (X - ScreenMinX) * Mod_TileEngine.TilePixelWidth + OffsetX

                    Call Mod_TileEngine.DrawGrh(.Grh, DrawableX, DrawableY, .Z, 1, 1, AlphaColor, , , True)
                End If
                .Alpha = .Alpha - (AlphaSwitch * .AlphaFactor)
            Else
                OverGrhTotalDisable = OverGrhTotalDisable + 1
            End If
        End With
    Next OverGrhIndex

    If OverGrhTotalDisable = MaxOverGrh Then
        Enable = False
    End If
End Sub

Private Sub Class_Initialize()
    Dim OverGrhIndex As Long
    For OverGrhIndex = 0 To MaxOverGrh - 1
        With OverGrhList(OverGrhIndex)
            .Alpha = 0
        End With
    Next OverGrhIndex
End Sub
Public Sub Add(ByVal X As Long, ByVal Y As Long, ByVal GrhIndex As Long, Optional ByVal Z As Double = 0#, _
               Optional ByVal FadeTime As Single = 1000, Optional ByVal Alpha As Double = 255)
    Enable = True

    Dim OverGrhIndex As Integer
    For OverGrhIndex = 0 To MaxOverGrh - 1
        With OverGrhList(OverGrhIndex)
            If .Alpha <= 0 Then
                .Alpha = Alpha
                .X = X
                .Y = Y
                .Z = Z
                .AlphaFactor = Alpha / (FadeTime / FadeStepTime)
                Call Mod_TileEngine.InitGrh(.Grh, GrhIndex)
                Exit Sub
            End If
        End With
    Next OverGrhIndex
End Sub
