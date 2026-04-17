Attribute VB_Name = "modSendData"
'**************************************************************
' SendData.bas - Has all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' Implemented by Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@gmail.com)
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

''
' Contains all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20070107

Option Explicit

Public Enum SendTarget
    ToUser
    ToAll
    toMap
    ToPCArea
    ToAllButIndex
    ToGM
    ToNPCArea
    ToNPCAreaButCounselors
    ToGuildMembers
    ToAdmins
    ToPCAreaButIndex
    ToAdminsAreaButConsejeros
    ToDiosesYclan
    ToConsejo
    ToClanArea
    ToConsejoCaos
    ToRolesMasters
    ToDeadArea
    ToPartyArea
    ToReal
    ToCaos
    ToRealYRMs
    ToCaosYRMs
    ToHigherAdmins
    ToGMsAreaButRmsOrCounselors
    ToUsersAreaButGMs
    ToUsersAndRmsAndCounselorsAreaButGMs
    ToAdminsButCounselorsAndRms
    ToHigherAdminsButRMs
    ToAdminsButCounselors
    ToRMsAndHigherAdmins
    ToAdminsButRMs
    ToDuelo
End Enum

Public Sub SendData(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndData As String, _
                    Optional ByVal IsDenounce As Boolean = False, Optional ByVal IsUrgent As Boolean = False)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus) - Rewrite of original
'Last Modify Date: 14/11/2010
'Last modified by: ZaMa
'14/11/2010: ZaMa - Now denounces can be desactivated.
'**************************************************************

On Error GoTo ErrHandler

    Dim LoopC As Long
    
    Select Case sndRoute
        Case SendTarget.ToUser
            If UserList(sndIndex).ConnIDValida Then
                TCP.Send UserList(sndIndex).Connection, IsUrgent
            End If
            
        Case SendTarget.ToPCArea
            Call SendToUserArea(sndIndex, sndData, IsUrgent)
            
        Case SendTarget.ToAdmins
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnIDValida Then
                    If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
                        ' Denounces can be desactivated
                        If IsDenounce Then
                            If UserList(LoopC).flags.SendDenounces Then
                                TCP.Send UserList(LoopC).Connection, IsUrgent
                            End If
                        Else
                            TCP.Send UserList(LoopC).Connection, IsUrgent
                        End If
                   End If
                End If
            Next LoopC

        Case SendTarget.ToAll
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnIDValida Then
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                        TCP.Send UserList(LoopC).Connection, IsUrgent
                    End If
                End If
            Next LoopC

        
        Case SendTarget.ToAllButIndex
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnIDValida) And (LoopC <> sndIndex) Then
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                        TCP.Send UserList(LoopC).Connection, IsUrgent
                    End If
                End If
            Next LoopC
        
        Case SendTarget.toMap
            Call SendToMap(sndIndex, sndData, IsUrgent)

        Case SendTarget.ToGuildMembers
            Call SendToGuildMembers(sndIndex, sndData, IsUrgent)

        Case SendTarget.ToDeadArea
            Call SendToDeadUserArea(sndIndex, sndData, IsUrgent)
        
        Case SendTarget.ToPCAreaButIndex
            Call SendToUserAreaButindex(sndIndex, sndData, IsUrgent)
        
        Case SendTarget.ToClanArea
            Call SendToUserGuildArea(sndIndex, sndData, IsUrgent)
        
        Case SendTarget.ToPartyArea
            Call SendToUserPartyArea(sndIndex, sndData, IsUrgent)

        Case SendTarget.ToAdminsAreaButConsejeros
            Call SendToAdminsButConsejerosArea(sndIndex, sndData, IsUrgent)

        Case SendTarget.ToNPCArea
            Call SendToNpcArea(sndIndex, sndData, IsUrgent)

        Case SendTarget.ToNPCAreaButCounselors
            Call SendToNpcAreaButCounselors(sndIndex, sndData, IsUrgent)
        
        Case SendTarget.ToDiosesYclan
            Call SendToDiosesYclan(sndIndex, sndData, IsUrgent)

        Case SendTarget.ToConsejo
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnIDValida) Then
                    If UserList(LoopC).flags.Privilegios And PlayerType.RoyalCouncil Then
                        TCP.Send UserList(LoopC).Connection, IsUrgent
                    End If
                End If
            Next LoopC

        Case SendTarget.ToConsejoCaos
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnIDValida) Then
                    If UserList(LoopC).flags.Privilegios And PlayerType.ChaosCouncil Then
                        TCP.Send UserList(LoopC).Connection, IsUrgent
                    End If
                End If
            Next LoopC

        Case SendTarget.ToRolesMasters
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnIDValida) Then
                    If UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster Then
                        TCP.Send UserList(LoopC).Connection, IsUrgent
                    End If
                End If
            Next LoopC

        Case SendTarget.ToReal
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnIDValida) Then
                    If UserList(LoopC).Faccion.ArmadaReal = 1 Or UserList(LoopC).Faccion.Alignment = eCharacterAlignment.FactionRoyal Then
                        TCP.Send UserList(LoopC).Connection, IsUrgent
                    End If
                End If
            Next LoopC
        
        Case SendTarget.ToCaos
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnIDValida) Then
                    If UserList(LoopC).Faccion.FuerzasCaos = 1 Or UserList(LoopC).Faccion.Alignment = eCharacterAlignment.FactionLegion Then
                        TCP.Send UserList(LoopC).Connection, IsUrgent
                    End If
                End If
            Next LoopC

        Case SendTarget.ToRealYRMs
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnIDValida) Then
                    If UserList(LoopC).Faccion.ArmadaReal = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        TCP.Send UserList(LoopC).Connection, IsUrgent
                    End If
                End If
            Next LoopC

        Case SendTarget.ToCaosYRMs
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnIDValida) Then
                    If UserList(LoopC).Faccion.FuerzasCaos = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        TCP.Send UserList(LoopC).Connection, IsUrgent
                    End If
                End If
            Next LoopC
 
        Case SendTarget.ToHigherAdmins
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnIDValida Then
                    If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
                        TCP.Send UserList(LoopC).Connection, IsUrgent
                   End If
                End If
            Next LoopC

        Case SendTarget.ToGMsAreaButRmsOrCounselors
            Call SendToGMsAreaButRmsOrCounselors(sndIndex, sndData, IsUrgent)

        Case SendTarget.ToUsersAreaButGMs
            Call SendToUsersAreaButGMs(sndIndex, sndData, IsUrgent)

        Case SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs
            Call SendToUsersAndRmsAndCounselorsAreaButGMs(sndIndex, sndData, IsUrgent)

        Case SendTarget.ToAdminsButCounselorsAndRms
            For LoopC = 1 To LastUser
                With UserList(LoopC)
                    If .ConnIDValida Then
                        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
                            If (.flags.Privilegios And (PlayerType.RoleMaster)) = 0 Then
                                TCP.Send UserList(LoopC).Connection, IsUrgent
                            End If
                       End If
                    End If
                End With
            Next LoopC

        Case SendTarget.ToHigherAdminsButRMs
            For LoopC = 1 To LastUser
                With UserList(LoopC)
                    If .ConnIDValida Then
                        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
                            If (.flags.Privilegios And (PlayerType.RoleMaster)) = 0 Then
                                TCP.Send UserList(LoopC).Connection, IsUrgent
                            End If
                       End If
                    End If
                End With
            Next LoopC
  
        Case SendTarget.ToAdminsButCounselors
            For LoopC = 1 To LastUser
                With UserList(LoopC)
                    If .ConnIDValida Then
                        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Or _
                            ((.flags.Privilegios And (PlayerType.RoleMaster)) <> 0 And (.flags.Privilegios And (PlayerType.Consejero)) <> 0) Then
                            TCP.Send UserList(LoopC).Connection, IsUrgent
                       End If
                    End If
                End With
            Next LoopC
  
        Case SendTarget.ToRMsAndHigherAdmins
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnIDValida) Then
                    If UserList(LoopC).flags.Privilegios And (PlayerType.RoleMaster Or PlayerType.Admin Or PlayerType.Dios) Then
                        TCP.Send UserList(LoopC).Connection, IsUrgent
                    End If
                End If
            Next LoopC

        Case SendTarget.ToAdminsButRMs
            For LoopC = 1 To LastUser
                With UserList(LoopC)
                    If .ConnIDValida Then
                        If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
                            If (.flags.Privilegios And (PlayerType.RoleMaster)) = 0 Then
                                TCP.Send UserList(LoopC).Connection, IsUrgent
                            End If
                       End If
                    End If
                End With
            Next LoopC

        Case SendTarget.ToDuelo
            If sndIndex > 0 Then
                For LoopC = 1 To LastUser
                    If (UserList(LoopC).ConnIDValida) Then
                        If UserList(LoopC).flags.DueloIndex = sndIndex Then
                            TCP.Send UserList(LoopC).Connection, IsUrgent
                        End If
                    End If
                Next LoopC
            End If

    End Select

ErrHandler:
    Protocol.Writer.Clear

    If Err.Number <> 0 Then
        Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendData de modSendData.bas")
    End If
End Sub

Private Sub SendToUserArea(ByVal UserIndex As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
    Dim query() As Collision.UUID
    Dim I       As Long
    
    For I = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
        Call TCP.Send(UserList(query(I).Name).Connection, IsUrgent)
    Next I
        
    Call TCP.Send(UserList(UserIndex).Connection, IsUrgent)
End Sub

Private Sub SendToUserAreaButindex(ByVal UserIndex As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
    Dim query() As Collision.UUID
    Dim I       As Long
    
    For I = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
        Call TCP.Send(UserList(query(I).Name).Connection, IsUrgent)
    Next I
End Sub

Private Sub SendToDeadUserArea(ByVal UserIndex As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
    Dim query() As Collision.UUID
    Dim I       As Long
    
    For I = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
        With UserList(query(I).Name)
            If (.flags.Muerto = 1 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) <> 0) Then
                Call TCP.Send(.Connection, IsUrgent)
            End If
        End With
    Next I

    Call TCP.Send(UserList(UserIndex).Connection, IsUrgent)
End Sub

Private Sub SendToUserGuildArea(ByVal UserIndex As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
    Dim query() As Collision.UUID
    Dim I       As Long
    
    For I = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
        With UserList(query(I).Name)
            If (.Guild.IdGuild = UserList(UserIndex).Guild.IdGuild Or ((.flags.Privilegios And PlayerType.Dios) And (.flags.Privilegios And PlayerType.RoleMaster) = 0)) Then
                Call TCP.Send(.Connection, IsUrgent)
            End If
        End With
    Next I

    Call TCP.Send(UserList(UserIndex).Connection, IsUrgent)
End Sub

Private Sub SendToUserPartyArea(ByVal UserIndex As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
    Dim query() As Collision.UUID
    Dim I       As Long
    
    For I = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
        With UserList(query(I).Name)
            If (.PartyIndex = UserList(UserIndex).PartyIndex) Then
                Call TCP.Send(.Connection, IsUrgent)
            End If
        End With
    Next I

    Call TCP.Send(UserList(UserIndex).Connection, IsUrgent)
End Sub

Private Sub SendToAdminsButConsejerosArea(ByVal UserIndex As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
    Dim query() As Collision.UUID
    Dim I       As Long
    
    For I = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
        With UserList(query(I).Name)
            If (.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin)) Then
                Call TCP.Send(.Connection, IsUrgent)
            End If
        End With
    Next I
    
    If (UserList(UserIndex).flags.Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin)) Then
        Call TCP.Send(UserList(UserIndex).Connection, IsUrgent)
    End If
End Sub

Private Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
    Dim query() As Collision.UUID
    Dim I       As Long
    
    For I = 0 To ModAreas.QueryObservers(NpcIndex, ENTITY_TYPE_NPC, query, ENTITY_TYPE_PLAYER)
        Call TCP.Send(UserList(query(I).Name).Connection, IsUrgent)
    Next I
End Sub

Private Sub SendToNpcAreaButCounselors(ByVal NpcIndex As Long, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
    Dim query() As Collision.UUID
    Dim I       As Long
    
    For I = 0 To ModAreas.QueryObservers(NpcIndex, ENTITY_TYPE_NPC, query, ENTITY_TYPE_PLAYER)
        If (UserList(query(I).Name).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero)) Then
            Call TCP.Send(UserList(query(I).Name).Connection, IsUrgent)
        End If
    Next I
End Sub

Public Sub SendToItemArea(ByVal Map As Integer, ByVal AreaX As Integer, ByVal AreaY As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
On Error GoTo ErrHandler

    Dim query() As Collision.UUID
    Dim I       As Long
    
    Dim ItemID  As Long
    ItemID = Pack(Map, AreaX, AreaY)
    
    For I = 0 To ModAreas.QueryObservers(ItemID, ENTITY_TYPE_OBJECT, query, ENTITY_TYPE_PLAYER)
        Call TCP.Send(UserList(query(I).Name).Connection, IsUrgent)
    Next I
    
ErrHandler:
    Protocol.Writer.Clear

    If Err.Number <> 0 Then
        Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendToItemArea de modSendData.bas")
    End If
End Sub

Public Sub SendToItemAreaButCounselors(ByVal Map As Integer, ByVal AreaX As Integer, ByVal AreaY As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
On Error GoTo ErrHandler

    Dim query() As Collision.UUID
    Dim I       As Long
        
    Dim ItemID  As Long
    ItemID = Pack(Map, AreaX, AreaY)
    
    For I = 0 To ModAreas.QueryObservers(ItemID, ENTITY_TYPE_OBJECT, query, ENTITY_TYPE_PLAYER)
        If (UserList(query(I).Name).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero)) Then
            Call TCP.Send(UserList(query(I).Name).Connection, IsUrgent)
        End If
    Next I
        
ErrHandler:
    Protocol.Writer.Clear

    If Err.Number <> 0 Then
        Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SendToItemArea de modSendData.bas")
    End If
End Sub

Private Sub SendToMap(ByVal Map As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)

    If MapInfo(Map).NumUsers < 1 Then Exit Sub

    Dim query() As Collision.UUID
    Dim I       As Long
    
    For I = 0 To ModAreas.QueryAt(Map, 50, 50, 50, query, ENTITY_TYPE_PLAYER)
        With UserList(query(I).Name)
            If ((.flags.Privilegios And Not PlayerType.User And Not PlayerType.Consejero And Not PlayerType.RoleMaster) = .flags.Privilegios) Then
                Call TCP.Send(.Connection, IsUrgent)
            End If
        End With
    Next I

End Sub

Private Sub SendToGMsAreaButRmsOrCounselors(ByVal UserIndex As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
    Dim query() As Collision.UUID
    Dim I       As Long
    
    For I = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
        With UserList(query(I).Name)
            If ((.flags.Privilegios And Not PlayerType.User And Not PlayerType.Consejero And Not PlayerType.RoleMaster) = .flags.Privilegios) Then
                Call TCP.Send(.Connection, IsUrgent)
            End If
        End With
    Next I
    
    With UserList(UserIndex)
        If ((.flags.Privilegios And Not PlayerType.User And Not PlayerType.Consejero And Not PlayerType.RoleMaster) = .flags.Privilegios) Then
            Call TCP.Send(.Connection, IsUrgent)
        End If
    End With
End Sub

Private Sub SendToUsersAreaButGMs(ByVal UserIndex As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
    Dim query() As Collision.UUID
    Dim I       As Long
    
    For I = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
        If (UserList(query(I).Name).flags.Privilegios And PlayerType.User) Then
            Call TCP.Send(UserList(query(I).Name).Connection, IsUrgent)
        End If
    Next I

    If (UserList(UserIndex).flags.Privilegios And PlayerType.User) Then
        Call TCP.Send(UserList(UserIndex).Connection, IsUrgent)
    End If
End Sub

Private Sub SendToUsersAndRmsAndCounselorsAreaButGMs(ByVal UserIndex As Integer, ByVal sdData As String, Optional ByVal IsUrgent As Boolean = False)
    Dim query() As Collision.UUID
    Dim I       As Long
    
    For I = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
        If (UserList(query(I).Name).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster)) Then
            Call TCP.Send(UserList(query(I).Name).Connection, IsUrgent)
        End If
    Next I

    If (UserList(UserIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster)) Then
        Call TCP.Send(UserList(UserIndex).Connection, IsUrgent)
    End If
End Sub

Public Sub AlertarFaccionarios(ByVal UserIndex As Integer)
      ' TODO
End Sub
