Attribute VB_Name = "Engine_Audio"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Argentum Online
'
' This program is free software; you can redistribute it and/or modify
' it under the terms of the Affero General Public License;
' either version 1 of the License, or any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' Affero General Public License for more details.
'
' You should have received a copy of the Affero General Public License
' along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''
' Module author: Agustin Alvarez <Wolftein> (18/8/2023)
'''

Option Explicit

Private Const EFFECT_FADE_TIME As Long = 2250 ' it takes 1.75sec to fade in/out between music

Private Const CHANNEL_MUSIC     As Long = 0
Private Const CHANNEL_EFFECT    As Long = 1
Private Const CHANNEL_INTERFACE As Long = 2

Private m_MasterMuted           As Boolean
Private m_MasterEnabled         As Boolean
Private m_MasterMuteOnFocusLost As Boolean
Private m_MasterVolume          As Long

Private m_MusicEnabled  As Boolean
Private m_MusicAsset    As Audio_Sound
Private m_MusicFilename As String
Private m_MusicID       As Long
Private m_MusicVolume   As Long

Private m_EffectEnabled As Boolean
Private m_EffectVolume  As Long

Private m_InterfaceEnabled As Boolean
Private m_InterfaceVolume  As Long

Private m_FadeAsset     As Audio_Sound
Private m_FadeFilename  As String
Private m_FadeID        As Long
Private m_FadeTick      As Long

Public Sub Initialize()
    
    m_MasterEnabled = True: m_MusicEnabled = True: m_EffectEnabled = True: m_InterfaceEnabled = True
    
End Sub

Public Sub Update(ByVal Tick As Long, ByVal CoordinatesX As Single, ByVal CoordinatesY As Single)
    
    Call Aurora_Audio.SetListener(CoordinatesX, 0#, CoordinatesY)
   
    ' Handle Enable/Disable on lost focus
    If (m_MasterMuteOnFocusLost And m_MasterEnabled) Then
        If (IsAppActive()) Then
            If (m_MasterMuted) Then
                Call Aurora_Audio.SetMasterVolume(m_MasterVolume * 0.01)
                
                m_MasterMuted = False
            End If
        Else
            If (Not m_MasterMuted) Then
                Call Aurora_Audio.SetMasterVolume(0)
                
                m_MasterMuted = True
            End If
        End If
    End If
    
    ' Handle Fade In/Out of the music channel
    If (m_FadeID <> 0) Then
        Call OnFadeEffect(Tick - m_FadeTick)
    End If

End Sub

Public Sub Halt()
    
    Call Aurora_Audio.Halt(CHANNEL_EFFECT)
    
End Sub

Public Function CreateEmitter(ByVal X As Single, ByVal Y As Single) As Audio_Emitter

    Set CreateEmitter = New Audio_Emitter
    
    With CreateEmitter
        Call .SetPosition(X, 0#, Y)
        Call .SetVelocity(0#, 0#, 0#)
        Call .SetInnerRadius(2#)
        Call .SetInnerRadiusAngle(3.14 / 4#)
        Call .SetAttenuation(12.25)
    End With

End Function

Public Sub UpdateEmitter(ByVal Emitter As Audio_Emitter, ByVal X As Single, ByVal Y As Single)

    If (Not Emitter Is Nothing) Then
        Call Emitter.SetPosition(X, 0#, Y)
    End If

End Sub

Public Sub DeleteEmitter(ByVal Emitter As Audio_Emitter, ByVal Immediately As Boolean)

    If (Not Emitter Is Nothing) Then
        Call Aurora_Audio.StopByEmitter(Emitter, Immediately)
    End If

    Set Emitter = Nothing
    
End Sub

Public Sub PlayMusic(ByVal fileName As String, Optional ByVal Repeat As Boolean = True, Optional ByVal Fade As Boolean = True)

    fileName = "Resources://MP3/" & fileName ' TODO: Wolftein -> should be normalized outside this function

    ' Check if we we're already playing the same song
    If (m_MusicFilename = fileName And m_MusicID <> 0 And m_FadeID = 0) Then Exit Sub
        
    ' Continue playing the music if we're trying to play the same music again while is active or fading
    If (Fade) Then
    
        ' If we were already fading, then exchange the fading song to master
        If (m_FadeID <> 0) Then
            If (m_FadeFilename = fileName) Then Exit Sub
            
            Call DisableSound(m_MusicID)
  
            m_MusicID = m_FadeID
            m_MusicFilename = m_FadeFilename
            Set m_MusicAsset = m_FadeAsset
        End If
        
        m_FadeFilename = fileName
        Set m_FadeAsset = Aurora_Content.Load(fileName, eResourceTypeSound)
        
        If (m_FadeAsset.GetStatus = eResourceStatusLoaded) Then
            m_FadeID = Aurora_Audio.Play(CHANNEL_MUSIC, m_FadeAsset, Nothing, Repeat)
            
            Call Aurora_Audio.SetGain(m_FadeID, 0#)
            Call Aurora_Audio.Start(m_FadeID)
        End If
        
        m_FadeTick = GetTickCount() ' TODO: Wolftein Change
    Else
        Call DisableMusic

        m_MusicFilename = fileName
        Set m_MusicAsset = Aurora_Content.Load(fileName, eResourceTypeSound)
        
        If (m_MusicAsset.GetStatus = eResourceStatusLoaded) Then
            m_MusicID = Aurora_Audio.Play(CHANNEL_MUSIC, m_MusicAsset, Nothing, Repeat)
            
            Call Aurora_Audio.Start(m_MusicID)
        End If
    End If
End Sub

Public Function PlayEffect(ByVal fileName As String, Optional ByVal Emitter As Audio_Emitter = Nothing, Optional ByVal Repeat As Boolean = False) As Long

    If Not m_EffectEnabled Then Exit Function

    PlayEffect = Play(CHANNEL_EFFECT, fileName, Emitter, Repeat)
    
End Function

Public Function PlayInterface(ByVal fileName As String) As Long

    If Not m_InterfaceEnabled Then Exit Function

    PlayInterface = Play(CHANNEL_INTERFACE, fileName, Nothing, False)
    
End Function

Public Sub DisableMusic()

    Call DisableSound(m_FadeID)
    Set m_FadeAsset = Nothing
    m_FadeFilename = vbNullString
    
    Call DisableSound(m_MusicID)
    Set m_MusicAsset = Nothing
    m_MusicFilename = vbNullString
    
End Sub

Public Sub DisableSound(ByRef InstanceID As Long, Optional ByVal Immediately As Boolean = True)

    If (InstanceID <> 0) Then
        Call Aurora_Audio.StopByID(InstanceID, Immediately)
    End If
    
    InstanceID = 0

End Sub

Public Property Let MasterMuteOnFocusLost(ByVal Activate As Boolean)

    m_MasterMuteOnFocusLost = Activate
    
End Property

Public Property Get MasterMuteOnFocusLost() As Boolean

    MasterMuteOnFocusLost = MasterMuteOnFocusLost
    
End Property

Public Property Get MasterEnabled() As Boolean

    MasterEnabled = m_MasterEnabled
    
End Property

Public Property Let MasterEnabled(ByVal Activate As Boolean)

    If m_MasterEnabled = Activate Then Exit Property

    m_MasterEnabled = Activate

    If Activate Then
        Call Aurora_Audio.SetMasterVolume(m_MasterVolume * 0.01)
    Else
        Call Aurora_Audio.SetMasterVolume(0)
    End If
    
End Property

Public Property Let MasterVolume(ByVal Volume As Long)

    If Volume < 0 Or Volume > 100 Then Exit Property

    Call Aurora_Audio.SetMasterVolume(Volume * 0.01)
    
    m_MasterVolume = Volume
    
End Property

Public Property Get MasterVolume() As Long

    MasterVolume = Aurora_Audio.GetMasterVolume() * 100
    
End Property

Public Property Get MusicEnabled() As Boolean

    MusicEnabled = m_MusicEnabled
    
End Property

Public Property Let MusicEnabled(ByVal Activate As Boolean)

    If m_MusicEnabled = Activate Then Exit Property

    m_MusicEnabled = Activate
    
    If Activate Then
        Call Aurora_Audio.SetSubmixVolume(CHANNEL_MUSIC, m_MusicVolume * 0.01)
        
        If (m_MusicID <> 0) Then
            Call Aurora_Audio.Start(m_MusicID)
        End If
    Else
        Call Aurora_Audio.SetSubmixVolume(CHANNEL_MUSIC, 0)
            
        If (m_MusicID <> 0) Then
            Call Aurora_Audio.Pause(m_MusicID)
        End If
    End If

End Property

Public Property Let MusicVolume(ByVal Volume As Long)

    If Volume < 0 Or Volume > 100 Then Exit Property

    Call Aurora_Audio.SetSubmixVolume(CHANNEL_MUSIC, Volume * 0.01)
    
    m_MusicVolume = Volume
    
End Property

Public Property Get MusicVolume() As Long

    MusicVolume = Aurora_Audio.GetSubmixVolume(CHANNEL_MUSIC) * 100
    
End Property

Public Property Get EffectEnabled() As Boolean

    EffectEnabled = m_EffectEnabled
    
End Property

Public Property Let EffectEnabled(ByVal Activate As Boolean)

    If m_EffectEnabled = Activate Then Exit Property

    m_EffectEnabled = Activate

    If Activate Then
        Call Aurora_Audio.SetSubmixVolume(CHANNEL_EFFECT, m_EffectVolume * 0.01)
    Else
        Call Aurora_Audio.SetSubmixVolume(CHANNEL_EFFECT, 0)
    End If
    
End Property

Public Property Let EffectVolume(ByVal Volume As Long)

    If Volume < 0 Or Volume > 100 Then Exit Property

    Call Aurora_Audio.SetSubmixVolume(CHANNEL_EFFECT, Volume * 0.01)

    m_EffectVolume = Volume
    
End Property

Public Property Get EffectVolume() As Long

    EffectVolume = Aurora_Audio.GetSubmixVolume(CHANNEL_EFFECT) * 100

End Property

Public Property Get InterfaceEnabled() As Boolean

    InterfaceEnabled = m_InterfaceEnabled
    
End Property

Public Property Let InterfaceEnabled(ByVal Activate As Boolean)

    If m_InterfaceEnabled = Activate Then Exit Property

    m_InterfaceEnabled = Activate

    If Activate Then
        Call Aurora_Audio.SetSubmixVolume(CHANNEL_INTERFACE, m_InterfaceVolume * 0.01)
    Else
        Call Aurora_Audio.SetSubmixVolume(CHANNEL_INTERFACE, 0)
    End If
    
End Property

Public Property Let InterfaceVolume(ByVal Volume As Long)

    If Volume < 0 Or Volume > 100 Then Exit Property

    Call Aurora_Audio.SetSubmixVolume(CHANNEL_INTERFACE, Volume * 0.01)

    m_InterfaceVolume = Volume
    
End Property

Public Property Get InterfaceVolume() As Long

    InterfaceVolume = Aurora_Audio.GetSubmixVolume(CHANNEL_INTERFACE) * 100
    
End Property

Private Function Play(ByVal Channel As Long, ByVal fileName As String, ByVal Emitter As Audio_Emitter, ByVal Repeat As Boolean) As Long

    fileName = "Resources://Wav/" & fileName ' TODO: Wolftein -> should be normalized outside this function

    Dim Effect As Audio_Sound
    Set Effect = Aurora_Content.Load(fileName, eResourceTypeSound)
    
    If (Effect.GetStatus = eResourceStatusLoaded) Then
        Play = Aurora_Audio.Play(Channel, Effect, Emitter, Repeat)
        
        If (Play <> 0) Then
            Call Aurora_Audio.Start(Play)
        End If
    End If
    
End Function

Private Sub OnFadeEffect(ByVal Delta As Long)

    Dim Time As Single
    Time = (Delta / EFFECT_FADE_TIME)
    
    ' Calculate the fade in/out gain between 0.0 and 1.0 using easeInOutCubic sigmoid function
    Dim Factor As Single
    Factor = IIf(Time < 0.5, 4 * Time * Time * Time, 1 - ((-2# * Time + 2#) ^ 3#) / 2#)
    
    ' Normalize in-case the user alt-tab and tick went to heaven (you will thank me later!)
    If (Factor > 1#) Then Factor = 1#
    
    ' Adjust both sound's gain value
    Call Aurora_Audio.SetGain(m_FadeID, Factor)
    
    If (m_MusicID <> 0) Then
        Call Aurora_Audio.SetGain(m_MusicID, 1# - Factor)
    End If
    
    ' Release the other sound if the fade has finished
    If (Delta >= EFFECT_FADE_TIME) Then
        Call DisableSound(m_MusicID, False)

        m_MusicID = m_FadeID
        m_MusicFilename = m_FadeFilename
        Set m_MusicAsset = m_FadeAsset
            
        m_FadeID = 0
        m_FadeFilename = vbNullString
        Set m_FadeAsset = Nothing
    End If
    
End Sub
