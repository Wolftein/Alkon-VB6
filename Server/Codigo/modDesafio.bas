Attribute VB_Name = "modDesafio"
'Argentum Online 0.14.0
'Copyright (C) 2002 Márquez Pablo Ignacio
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

'---------------------------------------------------------------------------------------
' Module    : modDesafio
' Date      : 12/01/2014
' Purpose   : Challenges between guilds
'---------------------------------------------------------------------------------------
Option Explicit

Public Type t_sand
    InUse As Byte
    
    IndexClan(1 To 2) As Byte
    DeadPoints(1 To 2) As Byte
    
    Amount_gold As Long
    Maxim_dead As Byte
    
    Event_time As Byte
    Time_start As Byte
    Event_map As Byte
    
    Invisibility As Byte
    Resucitar As Byte
    Elementary As Byte
    
    'Max_users As Byte
End Type

Public Const Maximus As Byte = 40
Public SandsChallenge(1 To Maximus) As t_sand

Public Type t_challenge
    'game
    InSand As Byte 'en que arena estoy jugando
    TeamSelect As Byte 'team 1 - team 2
    
    'pre game
    IndexOther As Integer
    ClanIndex As Integer
    Aceptar As Boolean
End Type

'***************************************************
'Autor: Mithrandir
'Last Modification: 23/01/2014
'***************************************************
Public Function Search_sand() As Byte

On Error GoTo ErrHandler
  
    Dim I As Long
    
    For I = 1 To Maximus
        If SandsChallenge(I).InUse = 0 Then
            Search_sand = I
            Exit Function
        End If
    Next I
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function Search_sand de modDesafio.bas")
End Function

'***************************************************
'Autor: Mithrandir
'Last Modification: 23/01/2014
'***************************************************
Public Sub Dead_challenge(ByVal Atacante As Integer, ByVal Victima As Integer)
        
On Error GoTo ErrHandler
  
    With UserList(Victima)
        'clansman ucciso
        If .Guild.IdGuild = UserList(Atacante).Guild.IdGuild Then Exit Sub
                
        'sono in diverse arene o non sono in alcun
        If .Challenge.InSand = 0 Or .Challenge.InSand <> UserList(Atacante).Challenge.InSand Then Exit Sub
        
        SandsChallenge(.Challenge.InSand).DeadPoints(.Challenge.TeamSelect) = SandsChallenge(.Challenge.InSand).DeadPoints(.Challenge.TeamSelect) + 1
        
        'es con alcance de muertes
        If SandsChallenge(.Challenge.InSand).Maxim_dead > 0 Then
            If SandsChallenge(.Challenge.InSand).DeadPoints(.Challenge.TeamSelect) = SandsChallenge(.Challenge.InSand).Maxim_dead Then
                Call Finish_challenge(Atacante)
                Exit Sub
            End If
        End If
        
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Dead_challenge de modDesafio.bas")
End Sub
'***************************************************
'Autor: Mithrandir
'Last Modification: 26/02/2014
'26/02/2014: Mithrandir - Agregado el limpiado de las variables de negociación
'***************************************************
Public Sub Start_challenge(ByVal UserIndex As Integer, ByVal Other_Acept As Integer, ByVal Arena As Byte)
    
On Error GoTo ErrHandler
  
    Dim I As Long
    
    'team 1
    With UserList(UserIndex)
        .Challenge.InSand = Arena
        .Challenge.TeamSelect = 1
        
        .Challenge.IndexOther = 0
        .Challenge.ClanIndex = 0
        
        SandsChallenge(Arena).IndexClan(1) = .Guild.IdGuild
        
        .Stats.GLD = .Stats.GLD - SandsChallenge(Arena).Amount_gold
        Call WriteUpdateGold(UserIndex)
        
        For I = 1 To LastUser
            If UserList(I).flags.UserLogged And UserList(I).Guild.IdGuild = .Guild.IdGuild Then
                UserList(I).Challenge.InSand = Arena
                UserList(I).Challenge.TeamSelect = 1
                
                'mensaje
                Call WriteConsoleMsg(I, "[DESAFIOS] El clan está en desafio de clanes.", FontTypeNames.FONTTYPE_CONSE)
                Call WriteUpdateChallengeStat(I)
            End If
        Next I
    End With
    
    'team 2
    With UserList(Other_Acept)
        .Challenge.InSand = Arena
        .Challenge.TeamSelect = 2
        
        .Challenge.IndexOther = 0
        .Challenge.ClanIndex = 0
        
        SandsChallenge(Arena).IndexClan(2) = .Guild.IdGuild
        
        .Stats.GLD = .Stats.GLD - SandsChallenge(Arena).Amount_gold
        Call WriteUpdateGold(UserIndex)
        
        For I = 1 To LastUser
            If UserList(I).flags.UserLogged And UserList(I).Guild.IdGuild = .Guild.IdGuild Then
                UserList(I).Challenge.InSand = Arena
                UserList(I).Challenge.TeamSelect = 2
                
                'mensaje
                Call WriteConsoleMsg(I, "[DESAFIOS] El clan aceptó un desafio de clanes.", FontTypeNames.FONTTYPE_CONSE)
                Call WriteUpdateChallengeStat(I)
            End If
        Next I
    End With
    
    'ocultamos formularios
    'Call WriteGuildWarCancel(UserIndex)
    'Call WriteGuildWarCancel(Other_Acept)
    Call WriteConsoleMsg(UserIndex, "WriteGuildWarCancel() no existe.", FontTypeNames.FONTTYPE_INFO)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Start_challenge de modDesafio.bas")
End Sub

'***************************************************
'Autor: Mithrandir
'Last Modification: 28/01/2014
'***************************************************
Public Sub Cancel_challenge(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
  
    Dim Other As Integer
    
    With UserList(UserIndex)
        .Challenge.ClanIndex = 0
        .Challenge.IndexOther = 0
        
        .Challenge.Aceptar = False
    End With

    'Call WriteGuildWarCancel(UserIndex)
    Call WriteConsoleMsg(UserIndex, "WriteGuildWar() no existe.", FontTypeNames.FONTTYPE_INFO)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Cancel_challenge de modDesafio.bas")
End Sub

'***************************************************
'Autor: Mithrandir
'Last Modification: 09/02/2014
'***************************************************
Public Sub Finish_challenge(ByVal UserIndex As Integer)
    Dim Num_Sand As Byte
    
On Error GoTo ErrHandler
  
    With UserList(UserIndex)
        Num_Sand = .Challenge.InSand
        
        .Stats.GLD = .Stats.GLD + (SandsChallenge(Num_Sand).Amount_gold * 2)
        SandsChallenge(Num_Sand).Amount_gold = 0
          
        Call SendData(SendTarget.ToDiosesYclan, .Guild.IdGuild, "[DESAFIOS] Tu clan ha ganado el desafio.")
                 
        'resetear all users with the same sand
        Dim I As Long
        
        For I = 1 To LastUser
            If UserList(I).Challenge.InSand = Num_Sand Then
                UserList(I).Challenge.InSand = 0
                UserList(I).Challenge.TeamSelect = 0
            End If
        Next I
        
        'reset values sand
        For I = 1 To 2
            SandsChallenge(Num_Sand).DeadPoints(I) = 0
            SandsChallenge(Num_Sand).IndexClan(I) = 0
        Next I
        
        SandsChallenge(Num_Sand).Elementary = 0
        SandsChallenge(Num_Sand).Invisibility = 0
        SandsChallenge(Num_Sand).Resucitar = 0
        
        SandsChallenge(Num_Sand).Event_map = 0
        SandsChallenge(Num_Sand).Event_time = 0
        SandsChallenge(Num_Sand).Time_start = 0
        
        SandsChallenge(Num_Sand).InUse = 0
        SandsChallenge(Num_Sand).Maxim_dead = 0
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Finish_challenge de modDesafio.bas")
End Sub
