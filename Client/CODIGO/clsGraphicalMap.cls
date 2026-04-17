VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicalMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsGraphicalMap
' Author    : Anagrama
' Date      : 12/08/2016
' Purpose   : Dibuja un mapa grafico e interactivo con todos sus chiches.
'---------------------------------------------------------------------------------------

Option Explicit

Private MapOffsetX As Long
Private MapOffsetY As Long

Private ownTilePixelWidth As Integer
Private ownTilePixelHeight As Integer

Private SelectedMap As Integer

Private VerNumeros As Byte
Private MapType As Byte

Private MapWidth As Byte
Private MapHeight As Byte

Private WithEvents MapWindow As PictureBox
Attribute MapWindow.VB_VarHelpID = -1

Private Device As Long

Public Sub Class_Terminate()
    Call Aurora_Graphics.DeletePass(Device)
End Sub
    
Public Property Get GetSelectedMap() As Integer
    GetSelectedMap = SelectedMap
End Property

Public Property Get GetMapType() As Byte
    GetMapType = MapType
End Property

Public Property Let SetMapType(ByVal TIPO As Byte)
    MapType = TIPO
    Call Invalidate(MapWindow.hwnd)
End Property

Public Property Get GetVerNumeros() As Byte
    GetVerNumeros = VerNumeros
End Property

Public Property Let SetVerNumeros(ByVal Ver As Byte)
    VerNumeros = Ver
    Call Invalidate(MapWindow.hwnd)
End Property

Public Sub DoScroll(ByVal Direccion As Byte)
On Error GoTo ErrHandler
  

    If CanScroll(Direccion) Then
        If Direccion = 1 Then
            MapOffsetY = MapOffsetY - 1
        ElseIf Direccion = 2 Then
            MapOffsetX = MapOffsetX + 1
        ElseIf Direccion = 3 Then
            MapOffsetY = MapOffsetY + 1
        ElseIf Direccion = 4 Then
            MapOffsetX = MapOffsetX - 1
        End If

        Call Invalidate(MapWindow.hwnd)
    End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoScroll de clsGraphicalMap.cls")
End Sub

Private Function CanScroll(ByVal Direccion As Byte)
On Error GoTo ErrHandler
  

    If Direccion = 1 Then
        CanScroll = MapOffsetY > 0
    ElseIf Direccion = 2 Then
        CanScroll = MapOffsetX < MapWidth - 9
    ElseIf Direccion = 3 Then
        CanScroll = MapOffsetY < MapHeight - 9
    ElseIf Direccion = 4 Then
        CanScroll = MapOffsetX > 0
    End If
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CanScroll de clsGraphicalMap.cls")
End Function

Private Function ClickMap(ByVal X As Long, ByVal Y As Long) As Long
On Error GoTo ErrHandler
  

    Dim Temp As Long
    Dim temp_x As Long
    Dim temp_y As Long
    
    temp_x = X \ ownTilePixelWidth + 1
    temp_y = Y \ ownTilePixelHeight + 1
    
    If temp_x > 9 Or temp_y > 9 Or temp_x < 1 Or temp_y < 1 Then Exit Function

    ClickMap = GraphicMap(temp_x + MapOffsetX, temp_y + MapOffsetY)
    
    If ClickMap > 0 Then
        Call ShowIntMapInfo(ClickMap)
    End If
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ClickMap de clsGraphicalMap.cls")
End Function

Friend Sub Initialize(ByRef MapPic As PictureBox, ByVal MyWidth As Byte, ByVal MyHeight As Byte, _
                            Optional ByVal FontSize As Integer = 7, Optional ByVal TileWidth As Integer = 64, _
                            Optional ByVal TileHeight As Integer = 64)
On Error GoTo ErrHandler
  

    Set MapWindow = MapPic

    'Make sure auto-redraw is set to false
    MapWindow.AutoRedraw = False
    
    'Set apropiate scale (pixel)
    MapWindow.ScaleMode = 3

    'Gives the width and height of the map for scrolling purposes.
    MapWidth = MyWidth
    MapHeight = MyHeight
    
    ownTilePixelHeight = TileHeight
    ownTilePixelWidth = TileWidth
    
    MapOffsetX = 0
    MapOffsetY = 0

    If (Device = 0) Then
        Device = Aurora_Graphics.CreatePassFromDisplay(MapPic.hwnd, MapPic.ScaleWidth, MapPic.ScaleHeight)
    End If
    
    Call CenterMap
    Call Invalidate(MapPic.hwnd)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Initialize de clsGraphicalMap.cls")
End Sub

Private Sub CenterMap()
On Error GoTo ErrHandler
  
    
    Dim X As Byte
    Dim Y As Byte
    
    For X = 1 To MapWidth
        For Y = 1 To MapHeight
            If GraphicMap(X, Y) = PlayerData.CurrentMap.Number Then
                If X > 5 Then
                    If X < MapWidth - 5 Then
                        MapOffsetX = X - 5
                    Else
                        MapOffsetX = MapWidth - 9
                    End If
                End If
                If Y > 5 Then
                    If Y < MapHeight - 5 Then
                        MapOffsetY = Y - 5
                    Else
                        MapOffsetY = MapHeight - 9
                    End If
                End If
                Exit For
            End If
        Next Y
    Next X
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CenterMap de clsGraphicalMap.cls")
End Sub

Private Sub MapWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim temp_x As Long
    Dim temp_y As Long

On Error GoTo ErrHandler
  
    If X < 0 Or Y < 0 Or X > MapWindow.Width Or Y > MapWindow.Height Then
        Exit Sub
    End If

    SelectedMap = ClickMap(X, Y)

    Call Invalidate(MapWindow.hwnd)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub MapWindow_MouseUp de clsGraphicalMap.cls")
End Sub

Public Sub MapWindow_Paint()
On Error GoTo ErrHandler

    Dim X As Long
    Dim Y As Long
    
    Call UIBegin(Device, MapWindow.ScaleWidth, MapWindow.ScaleHeight, &H0)

    For X = 1 To 9
        For Y = 1 To 9
            If GraphicMap(X + MapOffsetX, Y + MapOffsetY) > 0 Then

                Call DrawGrhIndex(IntMapInfo(GraphicMap(X + MapOffsetX, Y + MapOffsetY)).Grh, (X - 1) * ownTilePixelWidth, (Y - 1) * ownTilePixelHeight, -1, 0)

                If GraphicMap(X + MapOffsetX, Y + MapOffsetY) = PlayerData.CurrentMap.Number Then
                    Call DrawGrhIndex(4, (X - 1) * ownTilePixelWidth + 16, (Y - 1) * ownTilePixelHeight + 16, 0, 0)
                End If
                If VerNumeros Then
                    If GraphicMap(X + MapOffsetX, Y + MapOffsetY) <> SelectedMap Then
                        Call DrawText((X - 1) * ownTilePixelWidth + 20, (Y - 1) * ownTilePixelHeight + 20, 0#, CStr(GraphicMap(X + MapOffsetX, Y + MapOffsetY)), &HFFFFFF7D, eRendererAlignmentLeftTop, FuentesJuego.FuenteBase)
                    Else
                     Call DrawText((X - 1) * ownTilePixelWidth + 20, (Y - 1) * ownTilePixelHeight + 20, 0#, CStr(GraphicMap(X + MapOffsetX, Y + MapOffsetY)), &HFFFFFFFF, eRendererAlignmentLeftTop, FuentesJuego.FuenteBase)
                    End If
                End If
            End If
        Next Y
    Next X
    
    Call UIEnd
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RenderMap de clsGraphicalMap.cls")
End Sub


