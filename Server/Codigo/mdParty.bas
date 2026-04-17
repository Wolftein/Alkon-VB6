Attribute VB_Name = "mdParty"
'**************************************************************
' mdParty.bas - Library of functions to manipulate parties.
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

''
' SOPORTES PARA LAS PARTIES
' (Ver este modulo como una clase abstracta "PartyManager")
'


''
'cantidad maxima de parties en el servidor
Public Const MAX_PARTIES As Integer = 600

''
'Si esto esta en True, la exp sale por cada golpe que le da
'Si no, la exp la recibe al salirse de la party (pq las partys, floodean)
Public Const PARTY_EXPERIENCIAPORGOLPE As Boolean = False

''
'distancia al leader para que este acepte el ingreso
Public Const MAXDISTANCIAINGRESOPARTY As Byte = 2

''
'maxima distancia a un exito para obtener su experiencia
Public Const PARTY_MAXDISTANCIA As Byte = 18

''
'restan las muertes de los miembros?
Public Const CASTIGOS As Boolean = False

''
'Numero al que elevamos el nivel de cada miembro de la party
'Esto es usado para calcular la distribución de la experiencia entre los miembros
'Se lee del archivo de balance
Public ExponenteNivelParty As Single

''
'tPartyMember
'
' @param UserIndex UserIndex
' @param Experiencia Experiencia
'
Public Type tPartyMember
    UserIndex As Integer
    Experiencia As Double
End Type


Public Function NextParty() As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

Dim I As Integer
NextParty = -1
For I = 1 To MAX_PARTIES
    If Parties(I) Is Nothing Then
        NextParty = I
        Exit Function
    End If
Next I
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function NextParty de mdParty.bas")
End Function

Public Function PuedeCrearParty(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 11/04/2015 (D'Artagnan)
' - 05/22/2010 : staff members aren't allowed to party anyone. (Marco)
'11/04/2015: D'Artagnan - Skill points are no longer needed.
'***************************************************
On Error GoTo ErrHandler
  
    
    PuedeCrearParty = True
    
    If (UserList(UserIndex).flags.Privilegios And PlayerType.User) = 0 Then
    'staff members aren't allowed to party anyone.
        Call WriteConsoleMsg(UserIndex, "¡Los miembros del staff no pueden crear partys!", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
        PuedeCrearParty = False
        Exit Function
    'ElseIf CInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma)) * UserList(UserIndex).Stats.UserSkills(eSkill.Liderazgo) < 100 Then
    '    Call WriteConsoleMsg(UserIndex, "Tu carisma no es suficiente para liderar una party.", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
    '    PuedeCrearParty = False
    ElseIf UserList(UserIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_PARTY)
        PuedeCrearParty = False
        Exit Function
    End If
    
    If UserList(UserIndex).Stats.ELV < ConstantesBalance.MinCrearPartyLevel Then
        Call WriteConsoleMsg(UserIndex, "Tu nivel es muy bajo para crear o ingresar a una party.", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
        PuedeCrearParty = False
        Exit Function
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PuedeCrearParty de mdParty.bas")
End Function

Public Sub CrearParty(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 11/04/2015
'11/04/2015: D'Artagnan - Skill points are no longer needed.
'***************************************************
On Error GoTo ErrHandler
  

Dim tInt As Integer

With UserList(UserIndex)

    If UserList(UserIndex).Stats.ELV < ConstantesBalance.MinCrearPartyLevel Then
        Call WriteConsoleMsg(UserIndex, "Tu nivel es muy bajo para crear o ingresar a una party.", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
        Exit Sub
    End If

    If .PartyIndex = 0 Then
        If .flags.Muerto = 0 Then
            tInt = mdParty.NextParty
            If tInt = -1 Then
                Call WriteConsoleMsg(UserIndex, "Por el momento no se pueden crear más parties.", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
                Exit Sub
            Else
                Set Parties(tInt) = New clsParty
                If Not Parties(tInt).NuevoMiembro(UserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "La party está llena, no puedes entrar.", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
                    Set Parties(tInt) = Nothing
                    Exit Sub
                Else
                    Call WriteConsoleMsg(UserIndex, "¡Has formado una party!", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
                    .PartyIndex = tInt
                    .PartySolicitud = 0
                    If Not Parties(tInt).HacerLeader(UserIndex) Then
                        Call WriteConsoleMsg(UserIndex, "No puedes hacerte líder.", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
                    Else
                        Call WriteConsoleMsg(UserIndex, "¡Te has convertido en líder de la party!", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
                    End If
                End If
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_PARTY)
        End If
    Else
        Call WriteConsoleMsg(UserIndex, "Ya perteneces a una party.", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
    End If
End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CrearParty de mdParty.bas")
End Sub


Public Sub SalirDeParty(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

Dim PI As Integer
PI = UserList(UserIndex).PartyIndex
If PI > 0 Then
    If Parties(PI).SaleMiembro(UserIndex) Then
        'sale el leader
        Set Parties(PI) = Nothing
    Else
        UserList(UserIndex).PartyIndex = 0
    End If
Else
    Call WriteConsoleMsg(UserIndex, "No eres miembro de ninguna party.", FontTypeNames.FONTTYPE_INFO)
End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SalirDeParty de mdParty.bas")
End Sub

Public Sub ExpulsarDeParty(ByVal leader As Integer, ByVal OldMember As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

Dim PI As Integer
PI = UserList(leader).PartyIndex

If PI = UserList(OldMember).PartyIndex Then
    If Parties(PI).SaleMiembro(OldMember) Then
        'si la funcion me da true, entonces la party se disolvio
        'y los partyindex fueron reseteados a 0
        Set Parties(PI) = Nothing
    Else
        UserList(OldMember).PartyIndex = 0
    End If
Else
    Call WriteConsoleMsg(leader, LCase(UserList(OldMember).Name) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ExpulsarDeParty de mdParty.bas")
End Sub

''
' Determines if a user can use party commands like /acceptparty or not.
'
' @param User Specifies reference to user
' @return  True if the user can use party commands, false if not.
Public Function UserPuedeEjecutarComandos(ByVal User As Integer) As Boolean
'*************************************************
'Author: Marco Vanotti(Marco)
'Last modified: 05/05/09
'
'*************************************************
On Error GoTo ErrHandler
  
    Dim PI As Integer
    
    PI = UserList(User).PartyIndex
    
    If PI > 0 Then
        If Parties(PI).EsPartyLeader(User) Then
            UserPuedeEjecutarComandos = True
        Else
            Call WriteConsoleMsg(User, "¡No eres el líder de tu party!", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
            Exit Function
        End If
    Else
        Call WriteConsoleMsg(User, "No eres miembro de ninguna party.", FontTypeNames.FONTTYPE_INFO, eMessageType.Party)
        Exit Function
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function UserPuedeEjecutarComandos de mdParty.bas")
End Function

Public Sub BroadCastParty(ByVal UserIndex As Integer, ByRef texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

Dim PI As Integer
    
    PI = UserList(UserIndex).PartyIndex
    
    If PI > 0 Then
        Call Parties(PI).MandarMensajeAConsola(texto, UserList(UserIndex).Name)
    End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub BroadCastParty de mdParty.bas")
End Sub

Public Sub OnlineParty(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 11/27/09 (Budi)
'Adapte la función a los nuevos métodos de clsParty
'*************************************************
On Error GoTo ErrHandler
  
Dim I As Integer
Dim PI As Integer
Dim Text As String
Dim MembersOnline() As Integer
ReDim MembersOnline(1 To Constantes.MaxPartyMembers) As Integer

    PI = UserList(UserIndex).PartyIndex
    
    If PI > 0 Then
        Call Parties(PI).ObtenerMiembrosOnline(MembersOnline())
        Text = "Nombre(Exp): "
        For I = 1 To Constantes.MaxPartyMembers
            If MembersOnline(I) > 0 Then
                Text = Text & " - " & UserList(MembersOnline(I)).Name & " (" & Fix(Parties(PI).MiExperiencia(MembersOnline(I))) & ")"
            End If
        Next I
        Text = Text & ". Experiencia total: " & Parties(PI).ObtenerExperienciaTotal
        Call WriteConsoleMsg(UserIndex, Text, FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
    End If
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub OnlineParty de mdParty.bas")
End Sub


Public Sub TransformarEnLider(ByVal OldLeader As Integer, ByVal NewLeader As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

Dim PI As Integer

If OldLeader = NewLeader Then Exit Sub

PI = UserList(OldLeader).PartyIndex

If PI = UserList(NewLeader).PartyIndex Then
    If UserList(NewLeader).flags.Muerto = 0 Then
        If Parties(PI).HacerLeader(NewLeader) Then
            Call Parties(PI).MandarMensajeAConsola("El nuevo líder de la party es " & UserList(NewLeader).Name, UserList(OldLeader).Name)
        Else
            Call WriteConsoleMsg(OldLeader, "¡No se ha hecho el cambio de mando!", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
        End If
    Else
        Call WriteConsoleMsg(OldLeader, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(OldLeader, LCase$(UserList(NewLeader).Name) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub TransformarEnLider de mdParty.bas")
End Sub


Public Sub ActualizaExperiencias()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

'esta funcion se invoca antes de worlsaves, y apagar servidores
'en caso que la experiencia sea acumulada y no por golpe
'para que grabe los datos en los charfiles
Dim I As Integer

If Not PARTY_EXPERIENCIAPORGOLPE Then
    
    haciendoBK = True
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Distribuyendo experiencia en parties.", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
    For I = 1 To MAX_PARTIES
        If Not Parties(I) Is Nothing Then
            Call Parties(I).FlushExperiencia
        End If
    Next I
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Experiencia distribuida.", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    haciendoBK = False

End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ActualizaExperiencias de mdParty.bas")
End Sub

Public Sub ObtenerExito(ByVal UserIndex As Integer, ByVal Exp As Long, mapa As Integer, X As Integer, Y As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    If Exp <= 0 Then
        If Not CASTIGOS Then Exit Sub
    End If
    
    Call Parties(UserList(UserIndex).PartyIndex).ObtenerExito(Exp, mapa, X, Y)


  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ObtenerExito de mdParty.bas")
End Sub

Public Function CantMiembros(ByVal UserIndex As Integer) As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

CantMiembros = 0
If UserList(UserIndex).PartyIndex > 0 Then
    CantMiembros = Parties(UserList(UserIndex).PartyIndex).CantMiembros
End If

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CantMiembros de mdParty.bas")
End Function

''
' Sets the new p_sumaniveleselevados to the party.
'
' @param UserInidex Specifies reference to user
' @remarks When a user level up and he is in a party, we call this sub to don't desestabilice the party exp formula
Public Sub ActualizarSumaNivelesElevados(ByVal UserIndex As Integer)
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 28/10/08
'
'*************************************************
On Error GoTo ErrHandler
  
    If UserList(UserIndex).PartyIndex > 0 Then
        Call Parties(UserList(UserIndex).PartyIndex).UpdateSumaNivelesElevados(UserList(UserIndex).Stats.ELV)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ActualizarSumaNivelesElevados de mdParty.bas")
End Sub

Public Sub JoinMemberToParty(ByVal leader As Integer, ByVal NewMember As Integer)

On Error GoTo ErrHandler
  
    Dim PI As Integer
    Dim Reason As String, ReasonNewMember As String
    
    PI = UserList(leader).PartyIndex
    
    With UserList(NewMember)
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(leader, "¡Está muerto, no puedes aceptar miembros en ese estado!", _
                    FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
            Exit Sub
        End If
        
        If .PartyIndex > 0 Then
            Call SendData(SendTarget.ToUser, NewMember, _
                PrepareMessageConsoleMsg("Ya te encuentras en un grupo, retirate del grupo para poder ingresar a uno nuevo.", _
                FontTypeNames.FONTTYPE_PARTY))
            Exit Sub
        End If
        
        If Not Parties(PI).PuedeEntrar(NewMember, Reason, ReasonNewMember) Then
            Call WriteConsoleMsg(leader, Reason, FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
            Call WriteConsoleMsg(NewMember, ReasonNewMember, FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
            Exit Sub
        End If
        
        If Parties(PI).NuevoMiembro(NewMember) Then
            Call Parties(PI).MandarMensajeAConsola(UserList(leader).Name & _
                " ha agregado a " & .Name & " al grupo.", "Servidor")
            .PartyIndex = PI
        Else
            Call SendData(SendTarget.ToUser, leader, _
                PrepareMessageConsoleMsg("No se puede agregar a " & .Name & " al grupo porque el mismo se encuentra lleno.", _
                FontTypeNames.FONTTYPE_PARTY))
        End If

   
    End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub JoinMemberToParty de mdParty.bas")
End Sub
