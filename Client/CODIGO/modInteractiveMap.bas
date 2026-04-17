Attribute VB_Name = "modInteractiveMap"
Option Explicit

Public MapaGrafico As clsGraphicalMap

Public Enum eZoneType
    Segura = 0
    Insegura = 1
End Enum

Public Type tIntMapInfo
    Npcs As String
    Entradas As String
    Nivel As String
    Region As String
    Zona As Byte
    Name As String
    Grh As Long
End Type

Public IntMapInfo() As tIntMapInfo

Public GraphicMap() As Integer
Public IntMapWidth As Byte
Public IntMapHeight As Byte

Public Sub InicializarMapa(ByVal TIPO As Byte)
On Error GoTo ErrHandler
  
    Dim X As Byte
    Dim Y As Byte
    Dim MapType As String
    
    If TIPO = 0 Then
        MapType = "General"
    Else
        MapType = "Dungeon"
    End If
    
    IntMapWidth = Val(GetVar(IniPath & "MapOrg.dat", MapType, "MapWidth"))
    IntMapHeight = Val(GetVar(IniPath & "MapOrg.dat", MapType, "MapHeight"))
    
    ReDim GraphicMap(1 To IntMapWidth, 1 To IntMapHeight) As Integer
    
    For X = 1 To IntMapWidth
        For Y = 1 To IntMapHeight
            GraphicMap(X, Y) = Val(GetVar(IniPath & "MapOrg.dat", MapType, X & "-" & Y))
        Next Y
    Next X
    
    Call MapaGrafico.Initialize(frmMapa.picMapa, IntMapWidth, IntMapHeight)
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub InicializarMapa de modInteractiveMap.bas")
End Sub

Public Sub LoadIntMapInfo()
On Error GoTo ErrHandler
  
    Dim NumMaps As Integer
    Dim I As Integer
    
    NumMaps = Val(GetVar(IniPath & "MapData.dat", "INIT", "Mapas"))

    ReDim IntMapInfo(1 To NumMaps) As tIntMapInfo
    
    For I = 1 To NumMaps
        IntMapInfo(I).Npcs = GetVar(IniPath & "MapData.dat", "Mapa" & I, "Npcs")
        IntMapInfo(I).Entradas = GetVar(IniPath & "MapData.dat", "Mapa" & I, "Entradas")
        IntMapInfo(I).Nivel = GetVar(IniPath & "MapData.dat", "Mapa" & I, "Nivel")
        IntMapInfo(I).Region = GetVar(IniPath & "MapData.dat", "Mapa" & I, "Region")
        IntMapInfo(I).Name = GetVar(IniPath & "MapData.dat", "Mapa" & I, "Name")
        IntMapInfo(I).Zona = Val(GetVar(IniPath & "MapData.dat", "Mapa" & I, "Zona"))
        IntMapInfo(I).Grh = Val(GetVar(IniPath & "MapData.dat", "Mapa" & I, "Grh"))
    Next I
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadIntMapInfo de modInteractiveMap.bas")
End Sub

Public Sub ShowIntMapInfo(ByVal Mapa As Integer)
On Error GoTo ErrHandler
  
    With frmMapa
        .lblMapa = Mapa
        .lblNivel = IntMapInfo(Mapa).Nivel
        .lblZona = IIf(IntMapInfo(Mapa).Zona = eZoneType.Segura, "Insegura", "Segura")
        .lblNombre = IntMapInfo(Mapa).Name
        .txtEntradas = IntMapInfo(Mapa).Entradas
        .txtNpcs = IntMapInfo(Mapa).Npcs
        .lblRegion = IntMapInfo(Mapa).Region
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ShowIntMapInfo de modInteractiveMap.bas")
End Sub
