VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SoundMapInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' SoundMapInfo.cls
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

Private Type p_tSoundMapInfo
    Cantidad As Integer
    SoundIndex() As Integer
    flags() As Long
    Probabilidad() As Single
End Type

Private Enum p_eSoundFlags
    ninguna = 0
    Lluvia = 1
End Enum

Private p_Mapas() As p_tSoundMapInfo


'sonidos conocidos, pasados a enum para intelisense
Public Enum e_SoundIndex
    MUERTE_HOMBRE = 11
    MUERTE_MUJER = 74
    MORFAR_MANZANA = 82
    SOUND_COMIDA = 7
    MUERTE_MUJER_AGUA = 211
    MUERTE_HOMBRE_AGUA = 212
End Enum

Private Sub Class_Initialize()
'armar el array
On Error GoTo ErrHandler
  
    ReDim p_Mapas(1 To NumMaps) As p_tSoundMapInfo
    Call LoadSoundMapInfo
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Class_Initialize de clsMapSoundManager.cls")
End Sub

Public Sub LoadSoundMapInfo()
On Error GoTo ErrHandler
  
    Dim I As Integer
    Dim J As Integer
    Dim Temps As String
    Dim MAPFILE As String
    
    MAPFILE = ServerConfiguration.ResourcesPaths.Maps & "MAPA"
    
    'Usage of Val() prevents errors when dats are corrputed or incomplete. All invalid values are assumed to be zero.
    
'TODO : Log the error in the dat for correction.
    For I = 1 To UBound(p_Mapas)
        Temps = GetVar(MAPFILE & I & ".dat", "SONIDOS", "Cantidad")
        
        If IsNumeric(Temps) Then
            p_Mapas(I).Cantidad = Val(Temps)
            
            ReDim p_Mapas(I).flags(1 To p_Mapas(I).Cantidad) As Long
            ReDim p_Mapas(I).Probabilidad(1 To p_Mapas(I).Cantidad) As Single
            ReDim p_Mapas(I).SoundIndex(1 To p_Mapas(I).Cantidad) As Integer
            
            For J = 1 To p_Mapas(I).Cantidad
                p_Mapas(I).flags(J) = Val(GetVar(MAPFILE & I & ".dat", "SONIDO" & J, "Flags"))
                p_Mapas(I).Probabilidad(J) = Val(GetVar(MAPFILE & I & ".dat", "SONIDO" & J, "Probabilidad"))
                p_Mapas(I).SoundIndex(J) = Val(GetVar(MAPFILE & I & ".dat", "SONIDO" & J, "Sonido"))
            Next J
        Else
            p_Mapas(I).Cantidad = 0
        End If
    Next I
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadSoundMapInfo de clsMapSoundManager.cls")
End Sub

Public Sub ReproducirSonidosDeMapas()
On Error GoTo ErrHandler
  
    Dim I As Long
    Dim SonidoMapa As Byte
    Dim posX As Byte
    Dim posY As Byte
    
    'Sounds are played at a random position
    posX = RandomNumber(XMinMapSize, XMaxMapSize)
    posY = RandomNumber(YMinMapSize, YMaxMapSize)
    
    For I = 1 To UBound(p_Mapas)
        If p_Mapas(I).Cantidad > 0 Then
            SonidoMapa = RandomNumber(1, p_Mapas(I).Cantidad)
            If RandomNumber(1, 100) <= p_Mapas(I).Probabilidad(SonidoMapa) Then
                'tocarlo
                If Lloviendo Then
                    If p_Mapas(I).flags(SonidoMapa) Xor p_eSoundFlags.Lluvia Then
                        Call modSendData.SendToItemArea(I, posX, posY, PrepareMessagePlayWave(p_Mapas(I).SoundIndex(SonidoMapa), posX, posY))
                    End If
                Else
                    If p_Mapas(I).flags(SonidoMapa) Xor p_eSoundFlags.ninguna Then
                        Call modSendData.SendToItemArea(I, posX, posY, PrepareMessagePlayWave(p_Mapas(I).SoundIndex(SonidoMapa), posX, posY))
                    End If
                End If
            End If
        End If
    Next I
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ReproducirSonidosDeMapas de clsMapSoundManager.cls")
End Sub

Public Sub ReproducirSonido(ByVal Destino As SendTarget, ByVal Index As Integer, ByVal SoundIndex As Integer)
On Error GoTo ErrHandler
  
    Call SendData(Destino, Index, PrepareMessagePlayWave(SoundIndex, UserList(Index).Pos.X, UserList(Index).Pos.Y, UserList(Index).Char.CharIndex))
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ReproducirSonido de clsMapSoundManager.cls")
End Sub
