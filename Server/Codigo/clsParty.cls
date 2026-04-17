VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsParty.cls
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

Private p_members() As tPartyMember
'miembros

Private p_expTotal As Long
'Estadistica :D

Private p_Fundador As Integer
'el creador

Private p_CantMiembros As Integer
'cantidad de miembros

Private p_SumaNivelesElevados As Single
'suma de todos los niveles elevados a la ExponenteNivelParty > Esta variable se usa para calcular la experiencia repartida en la Party.

'datos en los pjs: | indexParty(indice en p_members), partyLeader(userindex del lider) |

'Constructor de clase
Public Sub Class_Initialize()
'***************************************************
'Author: Unknown
'Last Modification: 07/04/08
'Last Modification By: Marco Vanotti (MarKoxX)
' - 09/29/07 p_SumaNiveles added (Tavo)
' - 07/04/08 p_SumaNiveles changed to p_SumaNivelesElevados (MarKoxX)
'***************************************************
On Error GoTo ErrHandler
  
    ReDim p_members(1 To Constantes.MaxPartyMembers) As tPartyMember
    p_expTotal = 0
    p_CantMiembros = 0
    p_SumaNivelesElevados = 0
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Class_Initialize de clsParty.cls")
End Sub

'Destructor de clase

Public Sub Class_Terminate()
On Error GoTo ErrHandler
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Class_Terminate de clsParty.cls")
End Sub

''
' Sets the new p_sumaniveleselevados to the party.
'
' @param lvl Specifies reference to user level
' @remarks When a user level up and he is in a party, we update p_sumaNivelesElavados so the formula still works.
Public Sub UpdateSumaNivelesElevados(ByVal Lvl As Integer)
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 11/24/09
'11/24/09: Pato - Change the exponent to a variable with the exponent
'*************************************************
On Error GoTo ErrHandler
  
    p_SumaNivelesElevados = p_SumaNivelesElevados - ((Lvl - 1) ^ ExponenteNivelParty) + Lvl ^ ExponenteNivelParty
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateSumaNivelesElevados de clsParty.cls")
End Sub

Public Function MiExperiencia(ByVal UserIndex As Integer) As Long
'***************************************************
'Author: Unknown
'Last Modification: 11/27/09
'Last Modification By: Budi
' - 09/29/07 Experience is round to the biggest number less than that number
' - 09/29/07 Now experience is a real-number
' - 11/27/09 Arreglé el Out of Range.
'***************************************************
'Me dice cuanta experiencia tengo colectada ya en la party
On Error GoTo ErrHandler
  
Dim I As Integer
I = 1

While I <= Constantes.MaxPartyMembers And p_members(I).UserIndex <> UserIndex
    I = I + 1
Wend

If I <= Constantes.MaxPartyMembers Then
    MiExperiencia = Fix(p_members(I).Experiencia)
Else        'esto no deberia pasar :p
    MiExperiencia = -1
End If

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function MiExperiencia de clsParty.cls")
End Function

Public Sub ObtenerExito(ByVal ExpGanada As Long, ByVal mapa As Integer, X As Integer, Y As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 07/04/08
'Last Modification By: Marco Vanotti (MarKoxX)
' - 09/29/07 New formula for calculating the experience point of each user
' - 09/29/07 Experience is round to the biggest number less than that number
' - 09/29/07 Now experience is a real-number
' - 04/04/08 Ahora antes de calcular la experiencia a X usuario se fija si ese usuario existe (MarKoxX)
' - 07/04/08 New formula to calculate Experience for each user. (MarKoxX)
'***************************************************
'Se produjo un evento que da experiencia en la wp referenciada
On Error GoTo ErrHandler
  
Dim I As Integer
Dim UI As Integer
Dim expThisUser As Double

p_expTotal = p_expTotal + ExpGanada

For I = 1 To Constantes.MaxPartyMembers
    UI = p_members(I).UserIndex
    If UI > 0 Then
        ' Formula: Exp* (Nivel ^ ExponenteNivelParty) / sumadeNivelesElevados
        expThisUser = CDbl(ExpGanada * (UserList(p_members(I).UserIndex).Stats.ELV ^ ExponenteNivelParty) / p_SumaNivelesElevados)
        
        If mapa = UserList(UI).Pos.Map And UserList(UI).flags.Muerto = 0 Then
            If Distance(UserList(UI).Pos.X, UserList(UI).Pos.Y, X, Y) <= PARTY_MAXDISTANCIA Then
                p_members(I).Experiencia = p_members(I).Experiencia + expThisUser
                If p_members(I).Experiencia < 0 Then
                    p_members(I).Experiencia = 0
                End If
                If PARTY_EXPERIENCIAPORGOLPE Then
                    UserList(UI).Stats.Exp = UserList(UI).Stats.Exp + Fix(expThisUser)
                    Call CheckUserLevel(UI)
                End If
            End If
        End If
    End If
Next I

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ObtenerExito de clsParty.cls")
End Sub

Public Sub MandarMensajeAConsola(ByVal texto As String, ByVal Sender As String)
'feo feo, muy feo acceder a senddata desde aca, pero BUEEEEEEEEEEE...
'CHOTS: Más que "FEO" diría "DESAGRADABLEMENTE FEO"
On Error GoTo ErrHandler
  

Dim I As Integer

For I = 1 To Constantes.MaxPartyMembers
    If p_members(I).UserIndex > 0 Then
        Call WriteConsoleMsg(p_members(I).UserIndex, " [" & Sender & "] " & texto, FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
    End If
Next I

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub MandarMensajeAConsola de clsParty.cls")
End Sub

Public Function EsPartyLeader(ByVal UserIndex As Integer) As Boolean
On Error GoTo ErrHandler
  
    EsPartyLeader = (UserIndex = p_Fundador)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function EsPartyLeader de clsParty.cls")
End Function

Public Function NuevoMiembro(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 07/04/08
'Last Modification By: Marco Vanotti (MarKoxX)
' - 09/29/07 There is no level prohibition (Tavo)
' - 07/04/08 Added const ExponenteNivelParty. (MarKoxX)
'***************************************************
On Error GoTo ErrHandler
  

Dim I As Integer
I = 1
While I <= Constantes.MaxPartyMembers And p_members(I).UserIndex > 0
    I = I + 1
Wend

If I <= Constantes.MaxPartyMembers Then
    p_members(I).Experiencia = 0
    p_members(I).UserIndex = UserIndex
    NuevoMiembro = True
    p_CantMiembros = p_CantMiembros + 1
    p_SumaNivelesElevados = p_SumaNivelesElevados + (UserList(UserIndex).Stats.ELV ^ ExponenteNivelParty)
Else
    NuevoMiembro = False
End If

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function NuevoMiembro de clsParty.cls")
End Function

Public Function SaleMiembro(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 07/04/08
'Last Modification By: Marco Vanotti (MarKoxX)
' - 09/29/07 Experience is round to the biggest number less than that number
' - 09/29/07 Now experience is a real-number (Tavo)
' - 07/04/08 Added const ExponenteNivelParty. (MarKoxX)
'11/03/2010: ZaMa - Ahora no le dice al lider que salio de su propia party, y optimice con with.
'***************************************************
'el valor de retorno representa si se disuelve la party
On Error GoTo ErrHandler
  
    Dim I As Integer
    Dim J As Integer
    Dim MemberIndex As Integer
    
    I = 1
    SaleMiembro = False
    While I <= Constantes.MaxPartyMembers And p_members(I).UserIndex <> UserIndex
        I = I + 1
    Wend
    
    If I = 1 Then
        'sale el founder, la party se disuelve
        SaleMiembro = True
        Call MandarMensajeAConsola("El líder disuelve la party.", "Servidor")
        
        For J = Constantes.MaxPartyMembers To 1 Step -1
            
            With p_members(J)
            
                If .UserIndex > 0 Then
                    
                    ' No envia el mensaje al lider
                    If J <> 1 Then
                        Call WriteConsoleMsg(.UserIndex, "Abandonas la party liderada por " & _
                            UserList(p_members(1).UserIndex).Name & ".", FontTypeNames.FONTTYPE_PARTY)
                    End If
                    
                    Call WriteConsoleMsg(.UserIndex, "Durante la misma has conseguido " & CStr(Fix(.Experiencia)) & " puntos de experiencia.", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
            
                    If Not PARTY_EXPERIENCIAPORGOLPE Then
                        UserList(.UserIndex).Stats.Exp = UserList(.UserIndex).Stats.Exp + Fix(.Experiencia)
                        Call CheckUserLevel(.UserIndex)
                    End If
                    
                    Call MandarMensajeAConsola(UserList(.UserIndex).Name & " abandona la party.", "Servidor")
                    
                    UserList(.UserIndex).PartyIndex = 0
                    p_CantMiembros = p_CantMiembros - 1
                    p_SumaNivelesElevados = p_SumaNivelesElevados - (UserList(UserIndex).Stats.ELV ^ ExponenteNivelParty)
                    .UserIndex = 0
                    .Experiencia = 0
                    
                End If
                
            End With
            
        Next J
    Else
        If I <= Constantes.MaxPartyMembers Then
            
            MemberIndex = p_members(I).UserIndex
            
            With UserList(MemberIndex)
                If Not PARTY_EXPERIENCIAPORGOLPE Then
                    .Stats.Exp = .Stats.Exp + Fix(p_members(I).Experiencia)
                    Call CheckUserLevel(MemberIndex)
                End If
                
                Call MandarMensajeAConsola(.Name & " abandona la party.", "Servidor")
                'TODO: Revisar que esto este bien, y no este faltando/sobrando un mensaje, ahora solo los estoy corrigiendo
                Call WriteConsoleMsg(MemberIndex, "Durante la misma has conseguido " & _
                    CStr(Fix(p_members(I).Experiencia)) & " puntos de experiencia.", FontTypeNames.FONTTYPE_PARTY, eMessageType.Party)
                
                p_CantMiembros = p_CantMiembros - 1
                p_SumaNivelesElevados = p_SumaNivelesElevados - (UserList(UserIndex).Stats.ELV ^ ExponenteNivelParty)
                MemberIndex = 0
                p_members(I).Experiencia = 0
                p_members(I).UserIndex = 0
                CompactMemberList
            End With
        End If
    End If

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function SaleMiembro de clsParty.cls")
End Function

Public Function HacerLeader(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 09/29/07
'Last Modification By: Lucas Tavolaro Ortiz (Tavo)
' - 09/29/07 There is no level prohibition
'***************************************************
On Error GoTo ErrHandler
  
Dim I As Integer
Dim OldLeader As Integer
Dim oldExp As Double
Dim UserIndexIndex As Integer

UserIndexIndex = 0
HacerLeader = True

For I = 1 To Constantes.MaxPartyMembers
    If p_members(I).UserIndex > 0 Then
        If p_members(I).UserIndex = UserIndex Then
            UserIndexIndex = I
        End If
    End If
Next I

If Not HacerLeader Then Exit Function

If UserIndexIndex = 0 Then
    'catastrofe! esto no deberia pasar nunca! pero como es AO.... :p
    Call LogError("INCONSISTENCIA DE PARTIES")
    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("¡¡¡Inconsistencia de parties en HACERLEADER (UII = 0), AVISE A UN PROGRAMADOR ESTO ES UNA CATASTROFE!!!!", FontTypeNames.FONTTYPE_GUILD))
    HacerLeader = False
    Exit Function
End If


'aca esta todo bien y doy vuelta las collections
OldLeader = p_members(1).UserIndex
oldExp = p_members(1).Experiencia

p_members(1).UserIndex = p_members(UserIndexIndex).UserIndex     'que en realdiad es el userindex, pero no quiero inconsistencias moviendo experiencias
p_members(1).Experiencia = p_members(UserIndexIndex).Experiencia

p_members(UserIndexIndex).UserIndex = OldLeader
p_members(UserIndexIndex).Experiencia = oldExp

p_Fundador = p_members(1).UserIndex

'no need to compact
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function HacerLeader de clsParty.cls")
End Function


Public Sub ObtenerMiembrosOnline(ByRef MemberList() As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 09/29/07
'Last Modification By: Marco Vanotti (MarKoxX)
' - 09/29/07 Experience is round to the biggest number less than that number
' - 09/29/07 Now experience is a real-number (Tavo)
' - 08/18/08 Now TotalExperience is fixed (MarKoxX)
' - 11/27/09 Rehice la función, ahora devuelve el array con los UI online (Budi)
'***************************************************
On Error GoTo ErrHandler
  

Dim I As Integer

For I = 1 To Constantes.MaxPartyMembers
    If p_members(I).UserIndex > 0 Then
        MemberList(I) = p_members(I).UserIndex
    End If
Next I

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ObtenerMiembrosOnline de clsParty.cls")
End Sub

Public Function ObtenerExperienciaTotal() As Long
'***************************************************
'Author: Budi
'Last Modification: 11/27/09
'Retrieves the total experience acumulated in the party
'***************************************************
On Error GoTo ErrHandler
  
ObtenerExperienciaTotal = p_expTotal
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ObtenerExperienciaTotal de clsParty.cls")
End Function

Public Function PuedeEntrar(ByVal UserIndex As Integer, ByRef ReasonLeader As String, ByRef ReasonNewMember As String) As Boolean
On Error GoTo ErrHandler
  
    Dim esArmada As Boolean
    Dim esCaos As Boolean
    Dim MyLevel As Integer
    Dim I As Integer
    Dim rv As Boolean
    Dim UI As Integer
    
    rv = True
    esArmada = (UserList(UserIndex).Faccion.ArmadaReal = 1 Or (esArmada And UserList(UserIndex).Faccion.Alignment = eCharacterAlignment.FactionRoyal))
    esCaos = (UserList(UserIndex).Faccion.FuerzasCaos = 1 Or (esArmada And UserList(UserIndex).Faccion.Alignment = eCharacterAlignment.FactionLegion))
    MyLevel = UserList(UserIndex).Stats.ELV
    
    rv = Distancia(UserList(p_members(1).UserIndex).Pos, UserList(UserIndex).Pos) <= MAXDISTANCIAINGRESOPARTY

    ' Too far away from the target.
    If Not rv Then
        ReasonLeader = "El usuario " & UserList(UserIndex).Name & " se encuentra muy lejos."
        ReasonNewMember = "Te encuentras muy lejos del líder del grupo."
        PuedeEntrar = False
        Exit Function
    End If
    
    rv = (p_members(Constantes.MaxPartyMembers).UserIndex = 0)
    
    If Not rv Then
        ReasonLeader = "El grupo se encuentra lleno."
        ReasonNewMember = "El grupo se encuentra lleno."
        PuedeEntrar = False
        Exit Function
    End If
    
    For I = 1 To Constantes.MaxPartyMembers
        UI = p_members(I).UserIndex
        If UI > 0 Then
            If esArmada And UserList(UI).Faccion.Alignment = eCharacterAlignment.FactionLegion Then
                ReasonLeader = "Los miembros del Ejército Real no pueden entrar a una party con miembros de la Legión Oscura."
                ReasonNewMember = "Los miembros del Ejército Real no pueden entrar a una party con miembros de la Legión Oscura."
                rv = False
                Exit Function
            End If
            If esCaos And UserList(UI).Faccion.Alignment = eCharacterAlignment.FactionRoyal Then
                ReasonLeader = "Los miembros de la Legión Oscura no pueden entrar a una party con miembros del Ejército Real."
                ReasonNewMember = "Los miembros de la Legión Oscura no pueden entrar a una party con miembros del Ejército Real."
                rv = False
                Exit Function
            End If
                    
        End If
    Next I

    PuedeEntrar = rv

  
    Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PuedeEntrar de clsParty.cls")
End Function


Public Sub FlushExperiencia()
'***************************************************
'Author: Unknown
'Last Modification: 09/29/07
'Last Modification By: Lucas Tavolaro Ortiz (Tavo)
' - 09/29/07 Experience is round to the biggest number less than that number
' - 09/29/07 Now experience is a real-number
'***************************************************
'esta funcion se invoca frente a cerradas del servidor. Flushea la experiencia
'acumulada a los usuarios.
On Error GoTo ErrHandler
  

Dim I As Integer
If Not PARTY_EXPERIENCIAPORGOLPE Then   'esto sirve SOLO cuando acumulamos la experiencia!
    For I = 1 To Constantes.MaxPartyMembers
        If p_members(I).UserIndex > 0 Then
            If p_members(I).Experiencia > 0 Then
                UserList(p_members(I).UserIndex).Stats.Exp = UserList(p_members(I).UserIndex).Stats.Exp + Fix(p_members(I).Experiencia)
                Call CheckUserLevel(p_members(I).UserIndex)
            Else
                If Abs(UserList(p_members(I).UserIndex).Stats.Exp) > Abs(Fix(p_members(I).Experiencia)) Then
                    UserList(p_members(I).UserIndex).Stats.Exp = UserList(p_members(I).UserIndex).Stats.Exp + Fix(p_members(I).Experiencia)
                Else
                    UserList(p_members(I).UserIndex).Stats.Exp = 0
                End If
                
                Call WriteUpdateUserStats(p_members(I).UserIndex)
            End If
            
            p_members(I).Experiencia = 0
            
        End If
    Next I
End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub FlushExperiencia de clsParty.cls")
End Sub

Private Sub CompactMemberList()
On Error GoTo ErrHandler
  
Dim I As Integer
Dim freeIndex As Integer
I = 1
While I <= Constantes.MaxPartyMembers
    If p_members(I).UserIndex = 0 And freeIndex = 0 Then
        freeIndex = I
    ElseIf p_members(I).UserIndex > 0 And freeIndex > 0 Then
        p_members(freeIndex).Experiencia = p_members(I).Experiencia
        p_members(freeIndex).UserIndex = p_members(I).UserIndex
        p_members(I).UserIndex = 0
        p_members(I).Experiencia = 0
        'muevo el de la pos I a freeindex
        I = freeIndex
        freeIndex = 0
    End If
    I = I + 1
Wend

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CompactMemberList de clsParty.cls")
End Sub

Public Function CantMiembros() As Integer
On Error GoTo ErrHandler
  
CantMiembros = p_CantMiembros
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CantMiembros de clsParty.cls")
End Function
