Attribute VB_Name = "Mod_General"
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

Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public bFogata As Boolean

Public bLluvia() As Byte ' Array para determinar si

'debemos mostrar la animacion de la lluvia

Private lFrameTimer As Long

Public Function DirDats() As String
On Error GoTo ErrHandler

    DirDats = App.path & "\DAT\"
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function DirDats de General.bas")
End Function

Public Function DirGraficos() As String
On Error GoTo ErrHandler

    DirGraficos = App.path & "\" & GFX_PATH & "\"
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function DirGraficos de General.bas")
End Function

Public Function DirInterfaces() As String
On Error GoTo ErrHandler
  
    DirInterfaces = App.path & "\" & GFX_PATH & "\Interfaces\"
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function DirInterfaces de General.bas")
End Function

Public Function DirNpcMiniatures() As String
On Error GoTo ErrHandler
  
    DirNpcMiniatures = DirGraficos & "\Minis\"
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function DirInterfaces de General.bas")
End Function

Public Function DirSound() As String
On Error GoTo ErrHandler
  
    DirSound = App.path & "\" & SND_PATH & "\"
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function DirSound de General.bas")
End Function

Public Function DirMidi() As String
On Error GoTo ErrHandler
  
    DirMidi = App.path & "\" & MSC_PATH & "\"
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function DirMidi de General.bas")
End Function

Public Function DirMapas() As String
On Error GoTo ErrHandler
  
    DirMapas = App.path & "\" & MAP_PATH & "\"
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function DirMapas de General.bas")
End Function

Public Function DirExtras() As String
On Error GoTo ErrHandler
  
    DirExtras = App.path & "\EXTRAS\"
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function DirExtras de General.bas")
End Function

Public Function DirLogs() As String
On Error GoTo ErrHandler
  
    DirLogs = App.path & "\LOGS\"
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function DirLogs de General.bas")
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
On Error GoTo ErrHandler
  
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function RandomNumber de General.bas")
End Function

Public Function GetRawName(ByRef sName As String) As String
'***************************************************
'Author: ZaMa
'Last Modify Date: 13/01/2010
'Last Modified By: -
'Returns the char name without the clan name (if it has it).
'***************************************************
On Error GoTo ErrHandler
  

    Dim Pos As Integer
    
    Pos = InStr(1, sName, "<")
    
    If Pos > 0 Then
        GetRawName = Trim$(Left$(sName, Pos - 1))
    Else
        GetRawName = sName
    End If

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetRawName de General.bas")
End Function

Sub CargarAnimArmas()
On Error GoTo ErrHandler
  

    Dim loopc As Long
    Dim arch As String
    
    arch = App.path & "\init\" & "armas.dat"
    
    NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopc = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(GetVar(arch, "ARMA" & loopc, "Dir1")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(GetVar(arch, "ARMA" & loopc, "Dir2")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(GetVar(arch, "ARMA" & loopc, "Dir3")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(GetVar(arch, "ARMA" & loopc, "Dir4")), 0
    Next loopc
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CargarAnimArmas de General.bas")
End Sub
Sub CargarDialogos()
On Error GoTo ErrHandler
  
'***************************************************
'Author: Juan Dalmasso (CHOTS)
'Last Modify Date: 11/06/2011
'***************************************************
Dim archivoC As String
    
    archivoC = App.path & "\init\dialogos.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los diálogos. Falta el archivo dialogos.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim I As Byte
    
    For I = 1 To MAXCOLORESDIALOGOS
        ColoresDialogos(I).r = CByte(GetVar(archivoC, CStr(I), "R"))
        ColoresDialogos(I).g = CByte(GetVar(archivoC, CStr(I), "G"))
        ColoresDialogos(I).b = CByte(GetVar(archivoC, CStr(I), "B"))
    Next I
    

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CargarDialogos de General.bas")
End Sub
Sub CargarColores()
On Error GoTo ErrHandler
  
    Dim archivoC As String
    
    archivoC = App.path & "\init\colores.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim I As Long
    
    For I = 0 To 48 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(I).r = CByte(GetVar(archivoC, CStr(I), "R"))
        ColoresPJ(I).g = CByte(GetVar(archivoC, CStr(I), "G"))
        ColoresPJ(I).b = CByte(GetVar(archivoC, CStr(I), "B"))
    Next I
    
    ' Crimi
    ColoresPJ(50).r = CByte(GetVar(archivoC, "CR", "R"))
    ColoresPJ(50).g = CByte(GetVar(archivoC, "CR", "G"))
    ColoresPJ(50).b = CByte(GetVar(archivoC, "CR", "B"))
    
    ' Ciuda
    ColoresPJ(49).r = CByte(GetVar(archivoC, "CI", "R"))
    ColoresPJ(49).g = CByte(GetVar(archivoC, "CI", "G"))
    ColoresPJ(49).b = CByte(GetVar(archivoC, "CI", "B"))
    
    ' Atacable
    ColoresPJ(48).r = CByte(GetVar(archivoC, "AT", "R"))
    ColoresPJ(48).g = CByte(GetVar(archivoC, "AT", "G"))
    ColoresPJ(48).b = CByte(GetVar(archivoC, "AT", "B"))
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CargarColores de General.bas")
End Sub

Sub CargarAnimEscudos()
On Error GoTo ErrHandler

    Dim loopc As Long
    Dim arch As String
    
    arch = App.path & "\init\" & "escudos.dat"
    
    NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopc, "Dir1")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopc, "Dir2")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopc, "Dir3")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopc, "Dir4")), 0
    Next loopc
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CargarAnimEscudos de General.bas")
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, _
    Optional ByVal red As Integer = -1, Optional ByVal green As Integer, _
    Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, _
    Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = True, _
    Optional ByVal eeMessageType As eMessageType = eMessageType.None)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
'08/02/12 (D'Artagnan) - Multiple consoles.
'******************************************
On Error GoTo ErrHandler
  
    With RichTextBox
    
        If .Name <> frmMain.RecTxt(0).Name Then
            If Len(.Text) > 1000 Then
                'Get rid of first line
                .SelStart = InStr(1, .Text, vbCrLf) + 1
                .SelLength = Len(.Text) - .SelStart + 2
                .TextRTF = .SelRTF
            End If
                    
            .SelStart = Len(.Text)
            .SelLength = 0
            .SelBold = bold
            .SelItalic = italic
                    
            If Not red = -1 Then .SelColor = RGB(red, green, blue)
                    
            If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
            .SelText = Text
            .Refresh
        Else
            
            ' Always to General
            Call AddtoConsole(frmMain.RecTxt(eConsoleType.General), Text, red, green, blue, bold, italic, bCrLf)
            Call frmMain.NewMessageMainTab(eConsoleType.General)
            
            ' Determine which console to write in
            Dim vbConsoles(eConsoleType.Last - 1) As Boolean
                
            Select Case eeMessageType
                
                Case eMessageType.Info, eMessageType.m_MOTD
                    vbConsoles(eConsoleType.Custom) = (InfoMsg = 1)
                    
                Case eMessageType.Admin
                    vbConsoles(eConsoleType.Acciones) = True
                    vbConsoles(eConsoleType.Agrupaciones) = True
                    vbConsoles(eConsoleType.Custom) = (AdminMsg = 1)
                
                Case eMessageType.Party
                    vbConsoles(eConsoleType.Agrupaciones) = True
                    vbConsoles(eConsoleType.Custom) = (PartyMsg = 1)
                
                Case eMessageType.Guild
                    vbConsoles(eConsoleType.Agrupaciones) = True
                    vbConsoles(eConsoleType.Custom) = (GuildMsg = 1)
                    
                Case eMessageType.combate
                    vbConsoles(eConsoleType.Acciones) = True
                    vbConsoles(eConsoleType.Custom) = (CombateMsg = 1)
                    
                Case eMessageType.Trabajo
                    vbConsoles(eConsoleType.Acciones) = True
                    vbConsoles(eConsoleType.Custom) = (TrabajoMsg = 1)

            End Select
            
            Dim Index As Long
            For Index = 1 To eConsoleType.Last - 1 ' Ignore general
                If vbConsoles(Index) Then
                    Call AddtoConsole(frmMain.RecTxt(Index), Text, red, green, blue, bold, italic, bCrLf)
                    Call frmMain.NewMessageMainTab(Index)
             '       frmMain.RecTxt(Index).Refresh
                End If
            Next Index
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AddtoRichTextBox de General.bas")
End Sub

Sub AddtoConsole(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = True)
'******************************************
'Author: D'Artagnan
'Auxiliar sub for adding console messages
'******************************************
On Error GoTo ErrHandler
  
    With RichTextBox
        If Len(.Text) > 4000 Then
            'Get rid of half console
            .SelStart = InStr(2000, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        .SelText = Text
        
        RichTextBox.Refresh
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AddtoConsole de General.bas")
End Sub

Function AsciiValidos(ByRef cad As String) As Boolean
On Error GoTo ErrHandler
  
    Dim car As Byte
    Dim I As Long
    Dim J As Long
    
    cad = LCase$(cad)
    
    
    For I = 1 To Len(cad)
        AsciiValidos = False
        car = Asc(mid$(cad, I, 1))
        If ((car >= 97 And car <= 122)) Or (car = 32) Or (car = 46) Or (car = 44) Or (car = 59) Then
            AsciiValidos = True
        Else
            For J = 1 To Len(CAR_ESPECIALES)
                If car = Asc(mid$(CAR_ESPECIALES, J, 1)) Then
                    AsciiValidos = True
                    Exit For
                End If
            Next J
        End If
        If Not AsciiValidos Then Exit Function
    Next I
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function AsciiValidos de General.bas")
End Function

Function CheckUserData() As Boolean
    'Validamos los datos del user
On Error GoTo ErrHandler
  
    Dim loopc As Long
    Dim CharAscii As Integer
        
    If UserPassword = "" Then
        Call frmMessageBox.ShowMessage("Ingrese una contraseña")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, loopc, 1))
        If Not LegalCharacter(CharAscii, False) Then
            Call frmMessageBox.ShowMessage("Password inválido. El caracter " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    If UserName = "" Then
        Call frmMessageBox.ShowMessage("Ingrese un nombre de personaje.")
        Exit Function
    End If
        
    If Len(UserName) > MAX_NICKNAME_SIZE Then
        Call frmMessageBox.ShowMessage("El nombre de tu personaje no puede superar los " & MAX_NICKNAME_SIZE & " caracteres")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserName)
        CharAscii = Asc(mid$(UserName, loopc, 1))
        If Not LegalCharacter(CharAscii, True) Then
            Call frmMessageBox.ShowMessage("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    CheckUserData = True
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CheckUserData de General.bas")
End Function

Sub UnloadAllForms()
On Error GoTo ErrHandler
  

#If EnableSecurity Then
    Call UnprotectForm
#End If

    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UnloadAllForms de General.bas")
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer, ByVal inLogin As Boolean) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
On Error GoTo ErrHandler
  
Dim I As Long

    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        If inLogin Then 'Está chequeando un logueo
            For I = 1 To Len(CAR_ESPECIALES)
                If KeyAscii = Asc(mid$(CAR_ESPECIALES, I, 1)) Then
                    LegalCharacter = True
                    Exit Function
                End If
            Next I
            Exit Function
        Else
            Exit Function
        End If
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function LegalCharacter de General.bas")
End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
    'Set Connected
On Error GoTo ErrHandler

    Mod_Declaraciones.Connected = True

#If EnableSecurity Then
    'Unprotect character creation form
    Call UnprotectForm
#End If

    'Unload the connect form
    Unload frmCrearPersonaje
    Unload frmAccount
    Unload frmConnect
    
    frmMain.Second.Enabled = True
    frmMain.LblName.Caption = UserName
    'Load main form

    frmMain.Visible = True
    
    Call frmMain.ControlSM(eSMType.mSpells, False)
    Call frmMain.ControlSM(eSMType.mWork, False)
    Call frmMain.ControlSM(eSMType.mPets, True)
    
#If EnableSecurity Then
    'Protect the main form
    Call ProtectForm(frmMain)
#End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SetConnected de General.bas")
End Sub

Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/28/2008
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
' 06/28/2008: NicoNZ - Saqué lo que impedía que si el usuario estaba paralizado se ejecute el sub.
'***************************************************
On Error GoTo ErrHandler
  
    Dim LegalOk As Boolean

    If Cartel Then Cartel = False
    
    Select Case Direccion
        Case E_Heading.NORTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y - 1)
        Case E_Heading.EAST
            LegalOk = MoveToLegalPos(UserPos.X + 1, UserPos.Y)
        Case E_Heading.SOUTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y + 1)
        Case E_Heading.WEST
            LegalOk = MoveToLegalPos(UserPos.X - 1, UserPos.Y)
    End Select
    
    If LegalOk And Not UserParalizado Then
        
        If (UserMeditar) Then
            Call RequestMeditate
        Else
            Call WriteWalk(Direccion)
            
            MoveCharbyHead UserCharIndex, Direccion
            MoveScreen Direccion
        End If

    Else
        If charlist(UserCharIndex).Heading <> Direccion Then
            charlist(UserCharIndex).Heading = Direccion
                
            Call WriteChangeHeading(Direccion)
        End If
    End If
    
    If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub MoveTo de General.bas")
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************
On Error GoTo ErrHandler
  
    Call MoveTo(RandomNumber(NORTH, WEST))
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RandomMove de General.bas")
End Sub

Private Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
On Error GoTo ErrHandler

    Static LastTick As Single
    
    'No input allowed while Argentum is not the active window
    If Not Application.IsAppActive() Then Exit Sub
       
    ' Form main visible?
    If Not frmMain.Visible Then Exit Sub
       
    'No walking when in commerce or banking.
    If Comerciando Then Exit Sub
    
    If ViewingFormCantMove Then Exit Sub
    
    'No walking while writting in the forum.
    If MirandoForo Then Exit Sub
    
    'If game is paused, abort movement.
    If pausa Then Exit Sub
    
    'If looking at map, abort movement.
    If frmMapa.OnFocus = 1 Then Exit Sub
    
    'TODO: Debería informarle por consola?
    If Traveling Then Exit Sub
    
    
    LastTick = LastTick + GetInputElapsedTime()
    
    If (LastTick >= 16) Then

        ' This entire block will prevent the user from getting his position out of sync
        ' when meditating and moving.
        If (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
           (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0) Or _
           (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0) Or _
           (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0) Then
            
            ' Add pressed keys to array
            Call Movement.DirectionKeyDown(eKeyType.mKeyUp, NORTH)
            Call Movement.DirectionKeyDown(eKeyType.mKeyRight, EAST)
            Call Movement.DirectionKeyDown(eKeyType.mKeyDown, SOUTH)
            Call Movement.DirectionKeyDown(eKeyType.mKeyLeft, WEST)
                       
            If Not UserMoving And Not WaitInput Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                    If Not UserEstupido Then
                        ' Move to last pressed key
                        Call MoveTo(Movement.GetDirection())
                    Else
                        Call RandomMove
                    End If
                    
                    frmMain.Coord.Caption = PlayerData.CurrentMap.Number & " X: " & UserPos.X & " Y: " & UserPos.Y
                
            End If

            LastTick = 0
        End If
    End If

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CheckKeys de General.bas")
End Sub

Public Sub UpdateNodeScene(ByRef Node As Partitioner_Item, ByVal Id As Long, ByVal X As Long, ByVal Y As Long, ByVal Subtype As Long, ByVal Width As Long, ByVal Height As Long)
    
    Node.Id = Id
    Node.Type = Subtype
    Node.X = X
    Node.Y = Y
    Node.RectX1 = (X - Width / 2#)
    Node.RectY1 = (Y - Height)
    Node.RectX2 = Node.RectX1 + Width
    Node.RectY2 = Node.RectY1 + Height

End Sub

Public Sub UpdateNodeSceneChar(ByVal CharIndex As Long)
    Dim Width As Single, Height As Single
    Call GetCharacterDimension(CharIndex, Width, Height)

    With charlist(CharIndex)
        .Node.Id = UserCharIndex
        .Node.Type = 6
        .Node.X = .Pos.X
        .Node.Y = .Pos.Y
        .Node.RectX1 = (.Pos.X - Width / 2#)
        .Node.RectY1 = (.Pos.Y + IIf(.Nombre <> vbNullString, 1, 0) - Height)
        .Node.RectX2 = .Node.RectX1 + Width
        .Node.RectY2 = .Node.RectY1 + Height
    End With
End Sub


'TODO : Si bien nunca estuvo allí, el mapa es algo independiente o a lo sumo dependiente del engine, no va acá!!!
Public Function SwitchMap(ByVal Map As Integer) As Boolean
'**************************************************************
'Formato de mapas optimizado para reducir el espacio que ocupan.
'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
'**************************************************************
On Error GoTo ErrHandler
  
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte
        
        
    Dim Chunk As Memory_Chunk
    Set Chunk = Aurora_Content.Find("Resources://Mapas/Mapa" & Map & ".map")
    If (Not Chunk.HasData()) Then
        Exit Function
    End If
 
    Dim Reader As BinaryReader
    Set Reader = Chunk.GetReader()

    'map Header
    MapInfo.MapVersion = Reader.ReadInt16

    Call Reader.Skip(255 + 8 + 8) ' MiCabecera + Double
    
    Set Aurora_Scene = New Partitioner
    
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            With MapData(X, Y)
                
                ByFlags = Reader.ReadInt8
                
                .Blocked = (ByFlags And 1)
                
                .Graphic(1).GrhIndex = Reader.ReadInt16
                InitGrh .Graphic(1), .Graphic(1).GrhIndex
                    
                'Layer 2 used?
                If ByFlags And 2 Then
                    .Graphic(2).GrhIndex = Reader.ReadInt16
                    InitGrh .Graphic(2), .Graphic(2).GrhIndex
                  
                    Call InitGrhDepth(.Graphic(2), 2, X, Y, 0)
                    
                    With GrhData(.Graphic(2).GrhIndex)
                        Call UpdateNodeScene(MapData(X, Y).Nodes(2), -1, X, Y, 2, .TileWidth, .TileHeight)
                    End With
                    
                    
                    Call Aurora_Scene.Insert(.Nodes(2))
                    
                Else
                    .Graphic(2).GrhIndex = 0
                End If
                    
                'Layer 3 used?
                If ByFlags And 4 Then
                    .Graphic(3).GrhIndex = Reader.ReadInt16
                    InitGrh .Graphic(3), .Graphic(3).GrhIndex
                  
                    Call InitGrhDepth(.Graphic(3), 3, X, Y, 2)
                    
                    With GrhData(.Graphic(3).GrhIndex)
                        Call UpdateNodeScene(MapData(X, Y).Nodes(3), -1, X, Y, 3, .TileWidth, .TileHeight)
                    End With
                                      
                    Call Aurora_Scene.Insert(.Nodes(3))
                    
                Else
                    .Graphic(3).GrhIndex = 0
                End If
    
                'Layer 4 used?
                If ByFlags And 8 Then
                    .Graphic(4).GrhIndex = Reader.ReadInt16
                    InitGrh .Graphic(4), .Graphic(4).GrhIndex
                  
                    Call InitGrhDepth(.Graphic(4), 4, X, Y, 0)
                    
                    With GrhData(.Graphic(4).GrhIndex)
                        Call UpdateNodeScene(MapData(X, Y).Nodes(4), -1, X, Y, 4, .TileWidth, .TileHeight)
                    End With
                                                          
                    Call Aurora_Scene.Insert(.Nodes(4))
                    
                Else
                    .Graphic(4).GrhIndex = 0
                End If
                    
                'Trigger used?
                If ByFlags And 16 Then
                    .Trigger = Reader.ReadInt16
                Else
                    .Trigger = 0
                End If
                
                'Erase NPCs
                If .CharIndex > 0 Then
                    Call EraseChar(.CharIndex)
                End If
                
                'Erase OBJs
                If (.ObjGrh.GrhIndex > 0) Then
                    Call Engine_Audio.DeleteEmitter(.OBJInfo.SoundSource, True)
                    
                    .ObjGrh.GrhIndex = 0
                End If
                
            End With
        Next X
    Next Y

    MapInfo.Name = ""
    MapInfo.Music = ""
    
    CurMap = Map
    SwitchMap = True
    
    Exit Function
  
ErrHandler:
    
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SwitchMap de General.bas")
    SwitchMap = False
  
End Function

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
On Error GoTo ErrHandler
  
    Dim I As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For I = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next I
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ReadField de General.bas")
End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
On Error GoTo ErrHandler
  
    Dim Count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        Count = Count + 1
    Loop While curPos <> 0
    
    FieldCount = Count
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function FieldCount de General.bas")
End Function

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
On Error GoTo ErrHandler
  
    FileExist = (Dir$(File, FileType) <> "")
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function FileExist de General.bas")
End Function

Public Function IsIp(ByVal Ip As String) As Boolean
On Error GoTo ErrHandler
  
    Dim I As Long
    
    For I = 1 To UBound(ServersLst)
        If ServersLst(I).Ip = Ip Then
            IsIp = True
            Exit Function
        End If
    Next I
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function IsIp de General.bas")
End Function

Public Sub CargarServidores()
'********************************
'Author: Unknown
'Last Modification: 07/26/07
'Last Modified by: Rapsodius
'Added Instruction "CloseClient" before End so the mutex is cleared
'********************************
On Error GoTo errorH
    Dim f As String
    Dim c As Integer
    Dim I As Long
    
    f = App.path & "\init\sinfo.dat"
    c = Val(GetVar(f, "INIT", "Cant"))
    
    ReDim ServersLst(1 To c) As tServerInfo
    For I = 1 To c
        ServersLst(I).Desc = GetVar(f, "S" & I, "Desc")
        ServersLst(I).Ip = Trim$(GetVar(f, "S" & I, "Ip"))
        ServersLst(I).PassRecPort = CInt(GetVar(f, "S" & I, "P2"))
        ServersLst(I).Puerto = CInt(GetVar(f, "S" & I, "PJ"))
        ServersLst(I).PanelPassRecoveryUrl = GetVar(f, "S" & I, "PanelPassRecoveryUrl")
        
    Next I
        
    CurServer = 1
Exit Sub

errorH:
    Call MsgBox("Error cargando los servidores. " & "(" & Err.Description & ")", vbCritical + vbOKOnly, "Argentum Online")
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CargarServidores de General.bas")
    
    Call CloseClient
End Sub

Public Sub InitServersList()
On Error GoTo ErrHandler
  
    Dim I As Integer
    Dim Cont As Integer
    
    I = 1
    
    Do While (ReadField(I, RawServersList, Asc(";")) <> "")
        I = I + 1
        Cont = Cont + 1
    Loop
    
    ReDim ServersLst(1 To Cont) As tServerInfo
    
    For I = 1 To Cont
        Dim cur$
        cur$ = ReadField(I, RawServersList, Asc(";"))
        ServersLst(I).Ip = ReadField(1, cur$, Asc(":"))
        ServersLst(I).Puerto = ReadField(2, cur$, Asc(":"))
        ServersLst(I).Desc = ReadField(4, cur$, Asc(":"))
        ServersLst(I).PassRecPort = ReadField(3, cur$, Asc(":"))
    Next I
    
    CurServer = 1
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub InitServersList de General.bas")
End Sub

Public Function CurServerPasRecPort() As Integer
On Error GoTo ErrHandler
  
    If CurServer <> 0 Then
        CurServerPasRecPort = 7667
    Else
        CurServerPasRecPort = CInt(frmConnect.PortTxt)
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CurServerPasRecPort de General.bas")
End Function

Public Function CurServerIp() As String
On Error GoTo ErrHandler
  
    If CurServer <> 0 Then
        CurServerIp = ServersLst(CurServer).Ip
    Else
        CurServerIp = frmConnect.IPTxt.Text
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CurServerIp de General.bas")
End Function

Public Function CurServerPort() As Integer
On Error GoTo ErrHandler
  
    If CurServer <> 0 Then
        CurServerPort = ServersLst(CurServer).Puerto
    Else
        CurServerPort = CInt(frmConnect.PortTxt.Text)
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function CurServerPort de General.bas")
End Function

Sub Main()
' Compila sin ejecutar nada..
'On Error GoTo ErrHandler

#If MULTICLIENT = 0 Then
    If FindPreviousInstance Then
        Call MsgBox("Argentum Online ya está corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
        End
    End If
#End If

    'Read command line. Do it AFTER config file is loaded to prevent this from
    'canceling the effects of "/nores" option.
    Call LeerLineaComandos
    
    'usaremos esto para ayudar en los parches
    Call SaveSetting("ArgentumOnlineCliente", "Init", "Path", App.path & "\")
    
    ChDrive App.path
    ChDir App.path
    
    RandomClientToken = RandomString(32)

#If EnableSecurity And (Testeo = 0) Then
    'Obtener el HushMD5
    Dim fMD5HushYo As String * 32
    
    Set MD5 = New clsMD5

    fMD5HushYo = MD5.GetMD5File(App.path & "\" & App.EXEName & ".exe")

    Call MD5.MD5Reset
    MD5HushYo = txtOffset(hexMd52Asc(fMD5HushYo), 55)

#Else
    MD5HushYo = "0123456789AFSdef"  'We aren't using a real MD5
#End If

    'Init the protocol
#If EnableSecurity Then
    Call ProtocolPackets.InitProtocol
#Else
    Call Protocol.InitProtocol
#End If

    Call StartGame

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Main de General.bas")
End Sub

Public Sub StartGame()
    ' Load constants, classes, flags, graphics.
'On Error GoTo ErrHandler

    Call LoadInitialConfig
    

#If Testeo <> 1 Then
    'Dim PresPath As String
    'PresPath = DirGraficos & "Presentacion" & RandomNumber(1, 4) & ".jpg"
    
    'frmPres.Picture = LoadPicture(PresPath)
    'frmPres.Show vbModal    'Es modal, así que se detiene la ejecución de Main hasta que se desaparece
#End If

    
    frmConnect.Visible = True
    frmConnect.Show
    
    
    'Inicialización de variables globales
    ShowPerformanceData = False
    PrimeraVez = True
    prgRun = True
    pausa = False
    
    ' Start the interval manager for the connection actions
    Call MainTimer.SetInterval(TimersIndex.Action, INT_ACTION)
    Call MainTimer.Start(TimersIndex.Action)
    
    'Set the dialog's font
    'Dialogos.font = frmMain.font
    'DialogosClanes.font = frmMain.font
    
    lFrameTimer = GetTickCount

    Do While prgRun
        Call Aurora_Engine.Tick

        ' Update UI
        DoEvents
        
        'Sólo dibujamos si la ventana no está minimizada
        If frmMain.WindowState <> 1 And frmMain.Visible Then
            If (ShowNextFrame(frmMain.MouseX, frmMain.MouseY)) Then
                Mod_TileEngine.FPS = Mod_TileEngine.FPS + 1
            End If

            'Play ambient sounds
            Call RenderSounds
        Else
            Sleep 1
        End If

        Call CheckKeys

        'FPS Counter - mostramos las FPS
        If GetTickCount - lFrameTimer >= 1000 Then
            FramesPerSecCounter = FPS
            lFrameTimer = GetTickCount
            Mod_TileEngine.FPS = 0
        End If
        
#If EnableSecurity Then
        Call CheckSecurity
#End If

        ' Flush Visual Basic 6 - UI
        Call Aurora_Graphics.Flush

        ' Update audio thread
        Call Engine_Audio.Update(GetTickCount(), UserPos.X, UserPos.Y)
    Loop
    
    Call CloseClient
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub StartGame de General.bas")
End Sub

Private Sub LoadInitialConfig()
On Error GoTo ErrHandler
  
    frmCargando.Show
    frmCargando.Refresh
    

    '##############
    ' Aurora
    Call SetInitialConfigStatusIcon(frmCargando.imgMotorGrafico, False, "Aurora")
    
    Call LoadGameConfig
    Call LoadCustomMousePointers
    Call modCustomCursors.SetFormCursorDefault(frmCargando)
    
    If Not InitTileEngine(frmMain.hwnd, 32, 32, 13, 17, 8, 8) Then
        Call CloseClient
    End If
    Call SetInitialConfigStatusIcon(frmCargando.imgMotorGrafico, True, "Aurora")

    ' Load the game configuration from the UserConfig.ini file

    Call LoadClientSetup
    

    
    If GameConfig.Extras.bAskForResolutionChange Then
        Load frmResolution
        frmResolution.Show vbModal, frmCargando
        
        Call SaveGameConfig
    End If
    
    Call Resolution.SetResolution(GameConfig.Graphics.bUseFullScreen)

    '###########
    ' SERVIDORES
    'TODO : esto de ServerRecibidos no se podría sacar???
    Call CargarServidores
    ServersRecibidos = True

    '###########
    ' CONSTANTES
    Call SetInitialConfigStatusIcon(frmCargando.imgConstantes, False, "Constantes")
    
    Call InicializarNombres
    Call LoadEnabledClasses
    
    ' Initialize FONTTYPES
    Call InitFonts

    PlayerData.CurrentMap.Number = 1
    
    ' Mouse Pointer (Loaded before opening any form with buttons in it)
    If FileExist(DirExtras & "Hand.ico", vbArchive) Then _
        Set picMouseIcon = LoadPicture(DirExtras & "Hand.ico")
    
    Call SetInitialConfigStatusIcon(frmCargando.imgConstantes, True, "Constantes")
    '#######
    ' CLASES
    Call SetInitialConfigStatusIcon(frmCargando.imgClases, False, "Clases")
    
    Set Dialogos = New clsDialogs

    Set Inventario = New clsGraphicalInventory

    Set CustomKeys = New clsCustomKeys

    Set CustomMessages = New clsCustomMessages

    Set MainTimer = New clsTimer

    Set clsForos = New clsForum

    Set MD5 = New clsMD5

    Call SetInitialConfigStatusIcon(frmCargando.imgClases, True, "Clases")
    
#If EnableSecurity Then
    Call Security.StartSecurity
#End If

    '#######
    ' DATS
    Call SetInitialConfigStatusIcon(frmCargando.imgDats, False, "Dats")
    Call modQuests.LoadGuildQuests
    Call LoadNpcs
    Call LoadObjects

    Call SetInitialConfigStatusIcon(frmCargando.imgDats, True, "Dats")
    
    '###################
    ' ANIMACIONES EXTRAS
    Call SetInitialConfigStatusIcon(frmCargando.imgAnimaciones, False, "Animaciones")
    Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarColores
    Call CargarDialogos
    Call SetInitialConfigStatusIcon(frmCargando.imgAnimaciones, True, "Animaciones")
    
    '#############
    ' DIRECT SOUND
    Call SetInitialConfigStatusIcon(frmCargando.imgSonido, False, "Sonido")
    Call Engine_Audio.Initialize
    
    'Enable / Disable audio
    Engine_Audio.MasterVolume = GameConfig.Sounds.MasterVolume
    Engine_Audio.MusicVolume = GameConfig.Sounds.MusicVolume
    Engine_Audio.EffectVolume = GameConfig.Sounds.SoundsVolume
    Engine_Audio.InterfaceVolume = GameConfig.Sounds.InterfaceVolume
    Engine_Audio.MasterEnabled = GameConfig.Sounds.bMasterSoundEnabled
    Engine_Audio.MusicEnabled = GameConfig.Sounds.bMusicEnabled
    Engine_Audio.EffectEnabled = GameConfig.Sounds.bSoundEffectsEnabled
    Engine_Audio.InterfaceEnabled = GameConfig.Sounds.bInterfaceEnabled
    Engine_Audio.MasterMuteOnFocusLost = True ' TODO: Configuracion + Opciones
    
    Call Engine_Audio.PlayMusic(MP3_Inicio & ".mp3")
    
    Call SetInitialConfigStatusIcon(frmCargando.imgSonido, True, "Sonido")

    Call SetInitialConfigStatusIcon(frmCargando.imgExtras, False, "Extras")
    
    Call SetInitialConfigStatusIcon(frmCargando.imgExtras, True, "Extras")
    
    Call LoadIntMapInfo
    
    Call MessageManager.LoadFromIniFile(App.path & DAT_PATH & "Messages.dat")
    
    Call LoadGMCommand
    
    ' Inicializo aca para que se asigne bien el inventario
    Load frmMain
    Call frmMain.Inicializar

#If EnableSecurity Then
    CualMI = 0
    Call InitMI
#End If
    'Give the user enough time to read the welcome text
    Call Sleep(500)
    
    Unload frmCargando
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadInitialConfig de General.bas")
End Sub

Private Sub SetInitialConfigStatusIcon(ByRef Container As Image, ByVal Enabled As Boolean, ByRef loadingText As String)
On Error GoTo ErrHandler
  
    If Enabled Then
        Container.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "IconConfigEnabled.gif")
    Else
        Container.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "IconConfigDisabled.gif")
    End If
    
    frmCargando.lblStatus = loadingText
    DoEvents
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SetInitialConfigStatusIcon de General.bas")
End Sub

Public Sub LoadTimerIntervals()
'***************************************************
'Author: ZaMa
'Last Modification: 15/03/2011
'Set the intervals of timers
'***************************************************
On Error GoTo ErrHandler

    Call MainTimer.SetInterval(TimersIndex.Attack, PlayerData.Intervals.PlayerAttack)
    Call MainTimer.SetInterval(TimersIndex.Work, PlayerData.Intervals.Work)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithU, PlayerData.Intervals.UseItemWithKey)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, PlayerData.Intervals.UseItemDoubleClick)
    Call MainTimer.SetInterval(TimersIndex.SendRPU, PlayerData.Intervals.RequestPositionUpdate)
    Call MainTimer.SetInterval(TimersIndex.CastSpell, PlayerData.Intervals.PlayerCastSpell)
    Call MainTimer.SetInterval(TimersIndex.Arrows, PlayerData.Intervals.PlayerAttackArrow)
    Call MainTimer.SetInterval(TimersIndex.CastAttack, PlayerData.Intervals.PlayerAttackAfterSpell)
    Call MainTimer.SetInterval(TimersIndex.Meditate, PlayerData.Intervals.Meditate)
    Call MainTimer.SetInterval(TimersIndex.Click, 150) ' 150ms
        
    frmMain.macrotrabajo.Interval = PlayerData.Intervals.WorkMacro
    frmMain.macrotrabajo.Enabled = False
    
   'Init timers
    Call MainTimer.Start(TimersIndex.Attack)
    Call MainTimer.Start(TimersIndex.Work)
    Call MainTimer.Start(TimersIndex.UseItemWithU)
    Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
    Call MainTimer.Start(TimersIndex.SendRPU)
    Call MainTimer.Start(TimersIndex.CastSpell)
    Call MainTimer.Start(TimersIndex.Arrows)
    Call MainTimer.Start(TimersIndex.CastAttack)
    Call MainTimer.Start(TimersIndex.Meditate)
    Call MainTimer.Start(TimersIndex.Click)

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadTimerIntervals de General.bas")
End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
On Error GoTo ErrHandler
  
    WritePrivateProfileString Main, Var, value, File
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteVar de General.bas")
End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
On Error GoTo ErrHandler
  
    
    sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    GetPrivateProfileString Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetVar de General.bas")
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lo relativo a mapas, no tiene nada que hacer acá....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean
On Error GoTo ErrHandler
  
    HayAgua = ((MapData(X, Y).Graphic(1).GrhIndex >= 1505 And MapData(X, Y).Graphic(1).GrhIndex <= 1520) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 5665 And MapData(X, Y).Graphic(1).GrhIndex <= 5680) Or _
             MapData(X, Y).Graphic(1).GrhIndex >= 18974 And MapData(X, Y).Graphic(1).GrhIndex <= 18989 Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 13547 And MapData(X, Y).Graphic(1).GrhIndex <= 13562)) And _
                MapData(X, Y).Graphic(2).GrhIndex = 0
                
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function HayAgua de General.bas")
End Function

''
' Checks the command line parameters, if you are running Ao with /nores command and checks the AoUpdate parameters
'
'

Public Sub LeerLineaComandos()
'*************************************************
'Author: Unknown
'Last modified: 25/11/2008 (BrianPr)
'
'*************************************************
On Error GoTo ErrHandler
  
    Dim T() As String
    Dim I As Long
    
    Dim UpToDate As Boolean
    Dim TestEnvironment As Boolean
    
    'Parseo los comandos
    T = Split(Command, " ")
    For I = LBound(T) To UBound(T)
        Select Case UCase$(T(I))
            Case "/NORES" 'no cambiar la resolucion
                NoRes = True
            Case "/UPTODATE"
                UpToDate = True
            Case "/TEST" 'Use the test environment with the autoupdate.
                TestEnvironment = True
        End Select
    Next I

#If Testeo = 0 Then
    Call AoUpdate(UpToDate, NoRes, TestEnvironment)
#Else
    If App.LogMode = 0 Then NoRes = True
    If App.LogMode <> 0 Then Call AoUpdate(UpToDate, NoRes, TestEnvironment)
#End If

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LeerLineaComandos de General.bas")
End Sub

''
' Runs AoUpdate if we haven't updated yet, patches aoupdate and runs Client normally if we are updated.
'
' @param UpToDate Specifies if we have checked for updates or not
' @param NoREs Specifies if we have to set nores arg when running the client once again (if the AoUpdate is executed).

Private Sub AoUpdate(ByVal UpToDate As Boolean, ByVal NoRes As Boolean, ByVal TestEnvironment As Boolean)
'*************************************************
'Author: BrianPr
'Created: 25/11/2008
'Last modified: 25/11/2008
'
'*************************************************
On Error GoTo Error
    Dim extraArgs As String
    If Not UpToDate Then
        'No recibe update, ejecutar AU
        'Ejecuto el AoUpdate, sino me voy
        If Dir(App.path & "\AoUpdate.exe", vbArchive) = vbNullString Then
            MsgBox "No se encuentra el archivo de actualización AoUpdate.exe por favor descarguelo y vuelva a intentar", vbCritical
            End
        Else
            FileCopy App.path & "\AoUpdate.exe", App.path & "\AoUpdateTMP.exe"
            
            If NoRes Then
                extraArgs = " /nores"
            End If
            
            If TestEnvironment Then
                extraArgs = extraArgs & " /test"
            End If
            
            Call ShellExecute(0, "Open", App.path & "\AoUpdateTMP.exe", App.EXEName & ".exe" & extraArgs, App.path, SW_SHOWNORMAL)
            End
        End If
    Else
        If FileExist(App.path & "\AoUpdateTMP.exe", vbArchive) Then Kill App.path & "\AoUpdateTMP.exe"
    End If
Exit Sub

Error:
    If Err.Number = 75 Then 'Si el archivo AoUpdateTMP.exe está en uso, entonces esperamos 5 ms y volvemos a intentarlo hasta que nos deje.
        Sleep 5
        Resume
    Else
        MsgBox Err.Description & vbCrLf, vbInformation, "[ " & Err.Number & " ]" & " Error "
        End
    End If
End Sub

Private Sub LoadClientSetup()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/19/09
'11/19/09: Pato - Is optional show the frmGuildNews form
'**************************************************************
On Error GoTo ErrHandler
  
    Dim fHandle As Integer
          
    NoRes = Not GameConfig.Graphics.bUseFullScreen
    
    Set DialogosClanes = New clsGuildDlg
    DialogosClanes.Activo = Not GameConfig.Guilds.bShowDialogsInConsole
    DialogosClanes.CantidadDialogos = GameConfig.Guilds.MaxMessageQuantity

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadClientSetup de General.bas")
End Sub

Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************
On Error GoTo ErrHandler
  
    Ciudades(eCiudad.cUllathorpe) = "Ullathorpe"
    Ciudades(eCiudad.cNix) = "Nix"
    Ciudades(eCiudad.cBanderbill) = "Banderbill"
    Ciudades(eCiudad.cLindos) = "Lindos"
    Ciudades(eCiudad.cArghal) = "Arghâl"
    
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Elfo Oscuro"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"

    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladron"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Bandit) = "Bandido"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Worker) = "Trabajador"
    ListaClases(eClass.Hunter) = "Cazador"
  
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Evasión en combate"
    SkillsNames(eSkill.Armas) = "Combate cuerpo a cuerpo"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar árboles"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Minería"
    SkillsNames(eSkill.Carpinteria) = "Carpintería"
    SkillsNames(eSkill.Herreria) = "Herrería"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
    SkillsNames(eSkill.Wrestling) = "Combate sin armas"

    AtributosNames(eAtributos.Fuerza) = "Fuerza"
    AtributosNames(eAtributos.Agilidad) = "Agilidad"
    AtributosNames(eAtributos.Inteligencia) = "Inteligencia"
    AtributosNames(eAtributos.Carisma) = "Carisma"
    AtributosNames(eAtributos.Constitucion) = "Constitucion"
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub InicializarNombres de General.bas")
End Sub


''
' Removes all text from the console and dialogs

Public Sub CleanDialogs()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Removes all text from the console and dialogs
'**************************************************************
    'Clean console and dialogs
On Error GoTo ErrHandler
  
    Dim Index As Long
    For Index = 0 To eConsoleType.Last - 1
        frmMain.RecTxt(Index).Text = vbNullString
    Next Index
    
    Call DialogosClanes.RemoveDialogs
    
    Call Dialogos.RemoveAllDialogs
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CleanDialogs de General.bas")
End Sub

Public Sub CloseClient()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 8/14/2007
'Frees all used resources, cleans up and leaves
'**************************************************************
    ' Allow new instances of the client to be opened
On Error GoTo ErrHandler
  
    Call PrevInstance.ReleaseInstance
    
    EngineRun = False
    frmCargando.Show
    'Call AddtoRichTextBox(frmCargando.status, "Liberando recursos...", 0, 0, 0, 0, 0, 0)
    
    Call Resolution.ResetResolution
    
    'Stop tile engine
    Call DeinitTileEngine
    
    Call SaveGameConfig
    
    'Destruimos los objetos públicos creados
    Set CustomMessages = Nothing
    Set CustomKeys = Nothing

    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing

    Call UnloadAllForms

    End
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseClient de General.bas")
End Sub

Public Function esGM(CharIndex As Integer) As Boolean
On Error GoTo ErrHandler
  
esGM = False
If charlist(CharIndex).priv >= 1 And charlist(CharIndex).priv <= 5 Or charlist(CharIndex).priv = 25 Then _
    esGM = True

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function esGM de General.bas")
End Function

Public Function getTagPosition(ByVal Nick As String) As Integer
On Error GoTo ErrHandler
  
Dim Buf As Integer
Buf = InStr(Nick, "<")
If Buf > 0 Then
    getTagPosition = Buf
    Exit Function
End If
Buf = InStr(Nick, "[")
If Buf > 0 Then
    getTagPosition = Buf
    Exit Function
End If
getTagPosition = Len(Nick) + 2
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function getTagPosition de General.bas")
End Function

Public Function getStrenghtColor(ByVal iFuerza As Integer) As Long
On Error GoTo ErrHandler
  

    Dim m As Long
    Dim green As Byte
    Dim red As Byte
    
    m = 255 / MAXATRIBUTOS
    If m * iFuerza > 255 Then
        green = 255
        red = 255
    Else
        green = CByte(m * iFuerza)
    End If
    
    If green < 114 Then green = 114
    
    
    getStrenghtColor = RGB(red, green, 0)

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function getStrenghtColor de General.bas")
End Function

Public Function getDexterityColor(ByVal iAgilidad As Integer) As Long
On Error GoTo ErrHandler
  
    
    Dim m As Long
    Dim green As Byte
    m = 255 / MAXATRIBUTOS
    
    If m * iAgilidad > 255 Then green = 255 Else green = CByte(m * iAgilidad)
    
    If green < 114 Then green = 114
    
    getDexterityColor = RGB(255, green, 0)
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function getDexterityColor de General.bas")
End Function

Public Function EsAnuncio(ByVal ForumType As Byte) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 22/02/2010
'Returns true if the post is sticky.
'***************************************************
On Error GoTo ErrHandler
  
    Select Case ForumType
        Case eForumMsgType.ieCAOS_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieGENERAL_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieREAL_STICKY
            EsAnuncio = True
            
    End Select
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function EsAnuncio de General.bas")
End Function

Public Function ForumAlignment(ByVal yForumType As Byte) As Byte
'***************************************************
'Author: ZaMa
'Last Modification: 01/03/2010
'Returns the forum alignment.
'***************************************************
On Error GoTo ErrHandler
  
    Select Case yForumType
        Case eForumMsgType.ieCAOS, eForumMsgType.ieCAOS_STICKY
            ForumAlignment = eForumType.ieCAOS
            
        Case eForumMsgType.ieGeneral, eForumMsgType.ieGENERAL_STICKY
            ForumAlignment = eForumType.ieGeneral
            
        Case eForumMsgType.ieREAL, eForumMsgType.ieREAL_STICKY
            ForumAlignment = eForumType.ieREAL
            
    End Select
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ForumAlignment de General.bas")
End Function

Public Sub ResetAllInfo(ByVal CloseAll As Boolean)
'***************************************************
'Author: ZaMa
'Last Modification: 14/06/2011
'
'***************************************************
On Error GoTo ErrHandler
  
    Dim I As Long

    ' Disable timers
    frmMain.Second.Enabled = False
    frmMain.macrotrabajo.Enabled = False
    frmMain.tmrBlink.Enabled = False

    Mod_Declaraciones.Connected = False

    'Unload all forms except frmMain, frmConnect and frmCrearPersonaje
    If (CloseAll) Then
        Dim frm As Form
        For Each frm In Forms
            If frm.Name <> frmMain.Name And frm.Name <> frmConnect.Name And _
                frm.Name <> frmCrearPersonaje.Name And frm.Name <> frmAccount.Name And _
                frm.Name <> frmAccountChangePassword.Name And frm.Name <> frmMessageBox.Name And _
                frm.Name <> frmAccountCreate.Name Then
            
                Unload frm
            End If
        Next
    
        If Not frmAccount.Visible And Not frmConnect.Visible And _
           Not frmCrearPersonaje.Visible Then
            Call frmConnect.Show
        End If
    End If

    On Local Error GoTo 0
    
    ' Return to connection screen
    frmMain.Visible = False
    
    ' Return to inventory
    Call frmMain.ShowInventory
    
    'Stop audio
    If (CloseAll) Then
        Call Engine_Audio.Halt
    End If
    
    frmMain.IsPlaying = PlayLoop.plNone
    
    Call ResetGuildInfo
    Call GuildCleanInvitation
    
    ' Reset flags
    pausa = False
    UserMeditar = False
    UserEstupido = False
    UserCiego = False
    UserDescansar = False
    UserParalizado = False
    Traveling = False
    UserNavegando = False
    bRain = False
    bFogata = False
    Comerciando = False
    bShowTutorial = False
    RainBufferIndex = 0 '
    ViewingFormCantMove = False

    MirandoAsignarSkills = False
    MirandoCarpinteria = False
    MirandoEstadisticas = False
    MirandoForo = False
    MirandoHerreria = False
    MirandoParty = False
    
    '  related stuff
    If PetListQty Then
        For I = 1 To PetListQty
            PetList(I) = 0
        Next I
    End If
    PetListQty = 0
    PetSelectedIndex = 0
    HasPets = False
        
    'Delete all kind of dialogs
    Call CleanDialogs
    
#If EnableSecurity Then
    LOGGING = False
    LOGSTRING = False
    segLastPressed = 0
    LastMouse = False
    LastAmount = 0
#End If

    'Reset some char variables...

    For I = 1 To LastChar
#If EnableSecurity Then
        Call MI(CualMI).ResetInvisible(I)
#Else
        charlist(I).invisible = False
#End If
        charlist(I).IsSailing = False
    Next I

    ' Reset stats
    PlayerData.Class = 0
    PlayerData.Gender = 0
    PlayerData.Race = 0
    UserHogar = 0
    UserEmail = ""
    UserEstado = 0
    SkillPoints = 0
    Alocados = 0
    
    ' Reset skills
    For I = 1 To NUMSKILLS
        Set UserSkills(I) = New clsSkill
        Call UserSkills(I).Initialize
    Next I

    ' Reset attributes
    For I = 1 To NUMATRIBUTOS
        UserAtributos(I) = 0
    Next I

    ' Clear inventory slots
    Inventario.ClearAllSlots

#If EnableSecurity Then
    Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If

    Call PartyTempInviClear
    
    Call frmMain.hlst.Clean
    
    ' Connection screen midi
    Call Engine_Audio.PlayMusic(MP3_Inicio & ".mp3")

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetAllInfo de General.bas")
End Sub

Function complexNameToSimple(ByVal str As String, ByVal isGuild As Boolean) As String
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 06/12/2011
'
'***************************************************
On Error GoTo ErrHandler
  
Dim I As Long
Dim aux As String

    For I = 1 To Len(CAR_ESPECIALES)
    aux = mid$(CAR_ESPECIALES, I, 1)
        If InStr(1, str, aux) Then
            str = Replace(str, aux, mid$(CAR_COMUNES, I, 1))
        End If
    Next I
    
    If isGuild Then
    
        For I = 1 To Len(CAR_ESPECIALES_CLANES)
        aux = mid$(CAR_ESPECIALES_CLANES, I, 1)
            If InStr(1, str, aux) Then
                str = Replace(str, aux, mid$(CAR_COMUNES_CLANES, I, 1))
            End If
        Next I
    End If
    
    complexNameToSimple = str
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function complexNameToSimple de General.bas")
End Function

Public Sub LoadCustomConsole()
'***************************************************
'Author: D'Artagnan
'Last Modification: 01/27/2012
'
'***************************************************
On Error GoTo ErrHandler
  
    AdminMsg = 1
    GuildMsg = GetVar(App.path & CustomPath, "CONFIG", "Clan")
    PartyMsg = GetVar(App.path & CustomPath, "CONFIG", "Party")
    CombateMsg = GetVar(App.path & CustomPath, "CONFIG", "Combate")
    TrabajoMsg = GetVar(App.path & CustomPath, "CONFIG", "Trabajo")
    InfoMsg = GetVar(App.path & CustomPath, "CONFIG", "Info")

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadCustomConsole de General.bas")
End Sub

Public Function ValidInput(ParamArray Args() As Variant) As Boolean
'***************************************************
'Author: D'Artagnan
'Last Modification: 07/11/2014
'Return True if all provided strings are not empty.
'False otherwise.
'Use ValidInput(arg1, arg2, ...).
'***************************************************
On Error GoTo ErrHandler
  
    Dim I As Long
    
    For I = 0 To UBound(Args)
        If LenB(Args(I)) = 0 Then
            Exit Function
        End If
    Next I
    
    ValidInput = True
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ValidInput de General.bas")
End Function


Public Function IsSecondaryArmour(ByVal nObjectIndex As Integer) As Boolean
'***************************************************
'Author: D'Artagnan
'Last Modification: 03/02/2015
'Return True if the specified object index belongs
'to a secondary armour. False otherwise.
'***************************************************
On Error GoTo ErrHandler
  
    Dim sKey As String
    
    sKey = "OBJ" & CStr(nObjectIndex)
    IsSecondaryArmour = GameMetadata.Objs(nObjectIndex).Real = 2 Or _
                        GameMetadata.Objs(nObjectIndex).Caos = 2
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function IsSecondaryArmour de General.bas")
End Function

Public Function RandomString(cb As Integer) As String
On Error GoTo ErrHandler
  

    Randomize
    Dim rgch As String
    rgch = "abcdefghijklmnopqrstuvwxyz"
    rgch = rgch & UCase(rgch) & "0123456789"

    Dim I As Long
    For I = 1 To cb
        RandomString = RandomString & mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
    Next
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function RandomString de General.bas")
End Function

Public Sub LogError(ByVal ErrStr As String)
On Error GoTo ErrHandler
  
    Dim path As String
    Dim oFile As Integer
    
    Dim logsPath As String
    logsPath = DirLogs()
    
    ' Check for logs folder
    If Dir(logsPath, vbDirectory) = "" Then
        Call MkDir(logsPath)
    End If
    
    path = logsPath & "\Errores_" & Format(Now, "yyyyMMdd") & ".log"
    oFile = FreeFile
    
    Open path For Append As #oFile
        Print #oFile, Time & " - " & ErrStr
    Close #oFile
    
    Call OutputDebugString("AO: C - " & ErrStr)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LogError de General.bas")
End Sub

Public Sub ScreenCapture()
'**************************************************************
'Author: Unknown
'Last Modify Date: 11/16/2006
'**************************************************************
On Error GoTo Err:

    Dim dirFile As String
    
    ' Primero chequea si existe la carpeta Screenshots
    dirFile = App.path & "\Screenshots"
    If Not FileExist(dirFile, vbDirectory) Then Call MkDir(dirFile)

    ' TODO: Wolftein
    'Call wGL_Graphic.Capture(frmMain.hwnd, dirFile & "\" & Format(Now, "DD-MM-YYYY hh-mm-ss") & ".png")

    AddtoRichTextBox frmMain.RecTxt(0), "Screen Capturada!", 200, 200, 200, False, False, True, eMessageType.Info
Exit Sub

Err:
    Call AddtoRichTextBox(frmMain.RecTxt(0), Err.Number & "-" & Err.Description, 200, 200, 200, False, False, True, eMessageType.Info)

End Sub


Public Sub Invalidate(ByVal hwnd As Long)
    Dim udtRect As RECT

    Call GetClientRect(hwnd, udtRect)
    Call InvalidateRect(hwnd, udtRect, 1)
End Sub

Public Sub RequestMeditate()
    If (Not esGM(UserCharIndex)) Then
        WaitInput = True
    End If
    
    Call WriteMeditate
End Sub

Sub LoadMasteries()
On Error GoTo ErrHandler
  

    Dim I As Long
    Dim path As String
    Dim MasteriesQty As Integer
    
    path = App.path & "\dat\" & "Masteries.dat"

    GameMetadata.MasteriesQty = Val(GetVar(path, "INIT", "Masteries"))

    ReDim GameMetadata.Masteries(1 To GameMetadata.MasteriesQty) As tMetadataMastery
    For I = 1 To GameMetadata.MasteriesQty
        With GameMetadata.Masteries(I)
            .Id = I
            .Name = GetVar(path, "Mastery" & I, "Name")
            .Description = GetVar(path, "Mastery" & I, "Description")
            .Enabled = CBool(Val(GetVar(path, "Mastery" & I, "Enabled")))
            .IconGrh = Val(GetVar(path, "Mastery" & I, "IconGrh"))
            
            .RequiredGold = Val(GetVar(path, "Mastery" & I, "GoldRequired"))
            .RequiredPoints = Val(GetVar(path, "Mastery" & I, "PointsRequired"))
            .RequiredMastery = Val(GetVar(path, "Mastery" & I, "MasteryRequired"))
            
        End With
    Next I
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadMasteries de General.bas")
End Sub


Sub ShowLevelExpRequired()
    frmMain.lblPorcLvl.Caption = UserExp & "/" & UserPasarNivel
End Sub

Sub ShowLevelCompletionPerc()
    frmMain.lblPorcLvl.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"
End Sub


Public Sub PartyTempInviSave(ByVal UserNameRequest As String, ByVal UserIndexRequest As Integer)
    With PartyTempInvitation
        .UserNameRequest = UserNameRequest
        .UserIndexRequest = UserIndexRequest
    End With
End Sub

Public Sub PartyPendingInvitation()
    Dim Reply As Integer
    Dim IsAccepted As Boolean
    
    'TODO : sacar este msjbox  y agregar notificacion
    IsAccepted = False
    Reply = MsgBox(PartyTempInvitation.UserNameRequest & " te ha enviado una invitación para unirte a un grupo ", vbYesNo, "Invitacion de Grupo")
    
    If Reply = vbYes Then
        IsAccepted = True
    End If
    
    Call WritePartyAcceptInvitation(PartyTempInvitation.UserIndexRequest, PartyTempInvitation.UserNameRequest, IsAccepted)
    
    Exit Sub
End Sub
Public Sub PartyTempInviClear()
    PartyTempInvitation.UserIndexRequest = 0
    PartyTempInvitation.UserNameRequest = ""
End Sub

Public Function PartyInvitationEmpty() As Boolean
    PartyInvitationEmpty = (PartyTempInvitation.UserIndexRequest = 0) And (PartyTempInvitation.UserNameRequest = "")
End Function

Public Sub LoadObjects()
On Error GoTo ErrHandler
 
    Dim I As Long
    Dim path As String
    Dim IniReader As clsIniManager
    
    path = DirDats() & "OBJ.DAT"
    
    If Not FileExist(path, vbArchive) Then
        Call MsgBox("El archivo Obj.dat no existe")
        End
    End If
    
    Set IniReader = New clsIniManager
    Call IniReader.Initialize(path)
    
    GameMetadata.ObjsQty = IniReader.GetValueInt("INIT", "NumOBJs")
    
    If GameMetadata.ObjsQty <= 0 Then
        Call MsgBox("No hay objetos configurados en el archivo Obj.dat")
        End
    End If
    
    ReDim GameMetadata.Objs(1 To GameMetadata.ObjsQty) As tObjData
    
    For I = 1 To GameMetadata.ObjsQty
        With GameMetadata.Objs(I)
        
            .Name = IniReader.GetValue("OBJ" & I, "Name")
            .Real = Val(IniReader.GetValue("OBJ" & I, "Real"))
            .Caos = Val(IniReader.GetValue("OBJ" & I, "Caos"))
            .GrhIndex = Val(IniReader.GetValue("OBJ" & I, "GrhIndex"))
            .OBJType = Val(IniReader.GetValue("OBJ" & I, "ObjType"))
            .MinHit = Val(IniReader.GetValue("OBJ" & I, "MinHit"))
            .MaxHit = Val(IniReader.GetValue("OBJ" & I, "MaxHit"))
            .MinDef = Val(IniReader.GetValue("OBJ" & I, "MinDef"))
            .MaxDef = Val(IniReader.GetValue("OBJ" & I, "MaxDef"))
            .Valor = Val(IniReader.GetValue("OBJ" & I, "Valor"))
            .MinimumLevel = Val(IniReader.GetValue("OBJ" & I, "MinimumLevel"))
            
        End With
    Next I
    
    
    Set IniReader = Nothing
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadObjects de General.bas")
End Sub

Public Sub LoadNpcs()
On Error GoTo ErrHandler
 
    Dim I As Long
    Dim path As String
    Dim IniReader As clsIniManager
    
    path = DirDats() & "NPCs.DAT"
    
    If Not FileExist(path, vbArchive) Then
        Call MsgBox("El archivo NPCs.dat no existe")
        End
    End If
    
    Set IniReader = New clsIniManager
    Call IniReader.Initialize(path)
    
    GameMetadata.NpcsQty = IniReader.GetValueInt("INIT", "NumNPCs")
    
    If GameMetadata.NpcsQty <= 0 Then
        Call MsgBox("No hay NPCs configurados en el archivo NPCs.dat")
        End
    End If
    
    ReDim GameMetadata.Npcs(1 To GameMetadata.NpcsQty) As tNPC
    
    For I = 1 To GameMetadata.NpcsQty
        With GameMetadata.Npcs(I)
        
            .Name = IniReader.GetValue("NPC" & I, "Name")
            .MiniatureFileName = IniReader.GetValue("NPC" & I, "MiniatureFileName")

        End With
    Next I
    
    Set IniReader = Nothing
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadNpcs de General.bas")
End Sub

Public Sub LoadEnabledClasses()

    ListEnabledClasses(1) = eClass.Mage
    ListEnabledClasses(2) = eClass.Cleric
    ListEnabledClasses(3) = eClass.Warrior
    ListEnabledClasses(4) = eClass.Assasin
    ListEnabledClasses(5) = eClass.Thief
    ListEnabledClasses(6) = eClass.Bard
    ListEnabledClasses(7) = eClass.Druid
    ListEnabledClasses(8) = eClass.Paladin
    ListEnabledClasses(9) = eClass.Hunter

End Sub

Public Sub OpenDiscordLink()
On Error GoTo ErrHandler
    Call ShellExecute(0, "Open", "https://discord.gg/argentumonline", "", App.path, SW_SHOWNORMAL)
  
    Exit Sub
ErrHandler:
      Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub OpenDiscordLink de frmConnect.frm")
End Sub


''
' Initializes the fonts array

Public Sub InitFonts()

'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler

    With FontTypes(FontTypeNames.FONTTYPE_TALK)
        .red = 255
        .green = 255
        .blue = 255
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
        .red = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
        .red = 32
        .green = 51
        .blue = 223
        .bold = 1
        .italic = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        .red = 65
        .green = 190
        .blue = 156
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
        .red = 65
        .green = 190
        .blue = 156
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_EJECUCION)
        .red = 130
        .green = 130
        .blue = 130
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_PARTY)
        .red = 255
        .green = 180
        .blue = 250
    End With
    
    FontTypes(FontTypeNames.FONTTYPE_VENENO).green = 255
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILD)
        .red = 255
        .green = 255
        .blue = 255
        .bold = 1
    End With
    
    FontTypes(FontTypeNames.FONTTYPE_SERVER).green = 185
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
        .red = 228
        .green = 199
        .blue = 27
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJO)
        .red = 130
        .green = 130
        .blue = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOS)
        .red = 255
        .green = 60
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA)
        .green = 200
        .blue = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA)
        .red = 255
        .green = 50
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CENTINELA)
        .green = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GMMSG)
        .red = 255
        .green = 255
        .blue = 255
        .italic = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GM)
        .red = 30
        .green = 255
        .blue = 30
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_NEWBIE)
        .red = 242
        .green = 192
        .blue = 41
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_NEUTRAL)
        .red = 117
        .green = 117
        .blue = 117
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CITIZEN)
        .blue = 200
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSE)
        .red = 30
        .green = 150
        .blue = 30
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_DIOS)
        .red = 250
        .green = 250
        .blue = 150
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_NPCNAME)
        .red = 182
        .green = 169
        .blue = 81
        .bold = 1
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub InitFonts de Protocol.bas")
End Sub

