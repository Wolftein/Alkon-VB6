Attribute VB_Name = "ES"
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

Option Explicit

Public Sub CargarSpawnList()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  
    Dim Path As String
    Path = ServerConfiguration.ResourcesPaths.Dats
    Dim N As Integer, LoopC As Integer
    N = Val(GetVar(Path & "Invokar.dat", "INIT", "NumNPCs"))
    ReDim Declaraciones.SpawnList(N) As tCriaturasEntrenador
    For LoopC = 1 To N
        Declaraciones.SpawnList(LoopC).NpcIndex = Val(GetVar(Path & "Invokar.dat", "LIST", "NI" & LoopC))
        Declaraciones.SpawnList(LoopC).NpcName = GetVar(Path & "Invokar.dat", "LIST", "NN" & LoopC)
    Next LoopC
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CargarSpawnList de FileIO.bas")
End Sub

Function EsAdmin(ByRef Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 27/03/2011
'27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
'***************************************************
On Error GoTo ErrHandler
  
    EsAdmin = (Val(Administradores.GetValue("Admin", Name)) = 1)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function EsAdmin de FileIO.bas")
End Function

Function EsDios(ByRef Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 27/03/2011
'27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
'***************************************************
On Error GoTo ErrHandler
  
    EsDios = (Val(Administradores.GetValue("Dios", Name)) = 1)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function EsDios de FileIO.bas")
End Function

Function EsSemiDios(ByRef Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 27/03/2011
'27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
'***************************************************
On Error GoTo ErrHandler
  
    EsSemiDios = (Val(Administradores.GetValue("SemiDios", Name)) = 1)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function EsSemiDios de FileIO.bas")
End Function

Function EsGmEspecial(ByRef Name As String) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 27/03/2011
'27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
'***************************************************
On Error GoTo ErrHandler
  
    EsGmEspecial = (Val(Administradores.GetValue("Especial", Name)) = 1)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function EsGmEspecial de FileIO.bas")
End Function

Function EsConsejero(ByRef Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 27/03/2011
'27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
'***************************************************
On Error GoTo ErrHandler
  
    EsConsejero = (Val(Administradores.GetValue("Consejero", Name)) = 1)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function EsConsejero de FileIO.bas")
End Function

Function EsRolesMaster(ByRef Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 27/03/2011
'27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
'***************************************************
On Error GoTo ErrHandler
  
    EsRolesMaster = (Val(Administradores.GetValue("RM", Name)) = 1)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function EsRolesMaster de FileIO.bas")
End Function

Public Function EsGmChar(ByRef Name As String) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 27/03/2011
'Returns true if char is administrative user.
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim EsGm As Boolean
    
    ' Admin?
    EsGm = EsAdmin(Name)
    ' Dios?
    If Not EsGm Then EsGm = EsDios(Name)
    ' Semidios?
    If Not EsGm Then EsGm = EsSemiDios(Name)
    ' Consejero?
    If Not EsGm Then EsGm = EsConsejero(Name)

    EsGmChar = EsGm

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function EsGmChar de FileIO.bas")
End Function


Public Sub loadAdministrativeUsers()
'Admines     => Admin
'Dioses      => Dios
'SemiDioses  => SemiDios
'Especiales  => Especial
'Consejeros  => Consejero
'RoleMasters => RM
On Error GoTo ErrHandler
  

    'Si esta mierda tuviese array asociativos el código sería tan lindo.
    Dim Buf As Integer
    Dim I As Long
    Dim Name As String
       
    ' Public container
    Set Administradores = New clsIniManager
    
    ' Server ini info file
    Dim ServerIni As clsIniManager
    Set ServerIni = New clsIniManager
    
    Call ServerIni.Initialize(IniPath & "Server.ini")
    
       
    ' Admines
    Buf = Val(ServerIni.GetValue("INIT", "Admines"))
    
    For I = 1 To Buf
        Name = UCase$(ServerIni.GetValue("Admines", "Admin" & I))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Admin", Name, "1")

    Next I
    
    ' Dioses
    Buf = Val(ServerIni.GetValue("INIT", "Dioses"))
    
    For I = 1 To Buf
        Name = UCase$(ServerIni.GetValue("Dioses", "Dios" & I))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Dios", Name, "1")
        
    Next I
    
    ' Especiales
    Buf = Val(ServerIni.GetValue("INIT", "Especiales"))
    
    For I = 1 To Buf
        Name = UCase$(ServerIni.GetValue("Especiales", "Especial" & I))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Especial", Name, "1")
        
    Next I
    
    ' SemiDioses
    Buf = Val(ServerIni.GetValue("INIT", "SemiDioses"))
    
    For I = 1 To Buf
        Name = UCase$(ServerIni.GetValue("SemiDioses", "SemiDios" & I))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("SemiDios", Name, "1")
        
    Next I
    
    ' Consejeros
    Buf = Val(ServerIni.GetValue("INIT", "Consejeros"))
        
    For I = 1 To Buf
        Name = UCase$(ServerIni.GetValue("Consejeros", "Consejero" & I))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Consejero", Name, "1")
        
    Next I
    
    ' RolesMasters
    Buf = Val(ServerIni.GetValue("INIT", "RolesMasters"))
        
    For I = 1 To Buf
        Name = UCase$(ServerIni.GetValue("RolesMasters", "RM" & I))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("RM", Name, "1")
    Next I
    
    Set ServerIni = Nothing
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub loadAdministrativeUsers de FileIO.bas")
End Sub

Public Function GetCharPrivs(ByRef UserName As String) As PlayerType
'****************************************************
'Author: ZaMa
'Last Modification: 18/11/2010
'Reads the user's charfile and retrieves its privs.
'***************************************************
On Error GoTo ErrHandler
  

    Dim Privs As PlayerType

    If EsAdmin(UserName) Then
        Privs = PlayerType.Admin
        
    ElseIf EsDios(UserName) Then
        Privs = PlayerType.Dios

    ElseIf EsSemiDios(UserName) Then
        Privs = PlayerType.SemiDios
        
    ElseIf EsConsejero(UserName) Then
        Privs = PlayerType.Consejero
    
    Else
        Privs = PlayerType.User
    End If

    GetCharPrivs = Privs

  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetCharPrivs de FileIO.bas")
End Function

Public Function TxtDimension(ByVal Name As String) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim N As Integer, cad As String, Tam As Long
    N = FreeFile(1)
    Open Name For Input As #N
    Tam = 0
    Do While Not EOF(N)
        Tam = Tam + 1
        Line Input #N, cad
    Loop
    Close N
    TxtDimension = Tam
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function TxtDimension de FileIO.bas")
End Function

Public Sub CargarForbidenWords()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))
    Dim N As Integer, I As Integer
    N = FreeFile(1)
    Open DatPath & "NombresInvalidos.txt" For Input As #N
    
    For I = 1 To UBound(ForbidenNames)
        Line Input #N, ForbidenNames(I)
    Next I
    
    Close N

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CargarForbidenWords de FileIO.bas")
End Sub

Public Sub CargarHechizos()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'  ¡¡¡¡ NO USAR GetVar PARA LEER Hechizos.dat !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer Hechizos.dat se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

On Error GoTo ErrHandler

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."
    
    Dim Hechizo As Integer
    Dim I As Byte
    Dim TmpAreaEfficacy As String
    Dim TmpStr As String
    
    Dim Leer As clsIniManager
    Set Leer = New clsIniManager
    
    Call Leer.Initialize(DatPath & "Hechizos.dat")
    
    'obtiene el numero de hechizos
    NumeroHechizos = Val(Leer.GetValue("INIT", "NumeroHechizos"))
    
    ReDim Hechizos(1 To NumeroHechizos) As tHechizo
    
    frmCargando.Cargar.Min = 0
    frmCargando.Cargar.Max = NumeroHechizos
    frmCargando.Cargar.Value = 0
    
    'Llena la lista
    For Hechizo = 1 To NumeroHechizos
        With Hechizos(Hechizo)
            .Nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
            .Desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
            .PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
            
            .HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
            .TargetMsg = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
            .PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
            
            .Tipo = Val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
            .WAV = Val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
            .FXgrh = Val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
            
            .Loops = Val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
            
        '    .Resis = val(Leer.GetValue("Hechizo" & Hechizo, "Resis"))
            
            .SubeHP = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
            .MinHp = Val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
            .MaxHp = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
            
            .SubeMana = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
            .MinMana = Val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
            .MaxMana = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
            
            .SubeSta = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
            .MinSta = Val(Leer.GetValue("Hechizo" & Hechizo, "MinSta"))
            .MaxSta = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxSta"))
            
            .SubeHam = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeHam"))
            .MinHam = Val(Leer.GetValue("Hechizo" & Hechizo, "MinHam"))
            .MaxHam = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxHam"))
            
            .SubeSed = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeSed"))
            .MinSed = Val(Leer.GetValue("Hechizo" & Hechizo, "MinSed"))
            .MaxSed = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxSed"))
            
            .SubeAgilidad = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
            .MinAgilidad = Val(Leer.GetValue("Hechizo" & Hechizo, "MinAG"))
            .MaxAgilidad = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
            
            .SubeFuerza = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
            .MinFuerza = Val(Leer.GetValue("Hechizo" & Hechizo, "MinFU"))
            .MaxFuerza = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
            
            .SubeCarisma = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
            .MinCarisma = Val(Leer.GetValue("Hechizo" & Hechizo, "MinCA"))
            .MaxCarisma = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxCA"))
            
            .Invisibilidad = Val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
            .Paraliza = Val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
            .Inmoviliza = Val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
            .RemoverParalisis = Val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
            .RemoverEstupidez = Val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
            .RemueveInvisibilidadParcial = Val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
            
            .Putrefaccion = Val(Leer.GetValue("Hechizo" & Hechizo, "Putrefaccion"))
            .Teletransportacion = Val(Leer.GetValue("Hechizo" & Hechizo, "Teletransportacion"))
            .Salta = Val(Leer.GetValue("Hechizo" & Hechizo, "Salta"))
            .DistanciaSalto = Val(Leer.GetValue("Hechizo" & Hechizo, "DistanciaSalto"))
            .Petrificar = Val(Leer.GetValue("Hechizo" & Hechizo, "Petrificar"))
        
            .Area = Val(Leer.GetValue("Hechizo" & Hechizo, "Area"))
            If .Area > 0 Then
                ReDim .AreaEfficacy(1 To .Area)
                
                For I = 1 To .Area
                    TmpAreaEfficacy = Leer.GetValue("Hechizo" & Hechizo, "Area" & I)
                    .AreaEfficacy(I) = IIf(Len(TmpAreaEfficacy), Val(TmpAreaEfficacy), 100)
                Next I
            End If
            .CasterAffected = Leer.GetBooleanOrDefault("Hechizo" & Hechizo, "CasterAffected", True)
                        
            .Atraer = Val(Leer.GetValue("Hechizo" & Hechizo, "Atraer"))
            .ByPassPassive = Val(Leer.GetValue("Hechizo" & Hechizo, "ByPassPassive"))
            
            .CuraVeneno = Val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
            .Envenena = Val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
            .Maldicion = Val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
            .RemoverMaldicion = Val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
            .Bendicion = Val(Leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
            .Revivir = Val(Leer.GetValue("Hechizo" & Hechizo, "Revivir"))
            
            .Ceguera = Val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
            .Estupidez = Val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
            
            .Warp = Val(Leer.GetValue("Hechizo" & Hechizo, "Warp"))
            
            .Invoca = Val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
            .NumNpc = Val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
            .cant = Val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
            .Mimetiza = Val(Leer.GetValue("hechizo" & Hechizo, "Mimetiza"))
            
        '    .Materializa = val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
        '    .ItemIndex = val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
            
            .MinSkill = Val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
            
            If Hechizo = 38 Then
                Debug.Print
            End If
            
            .ManaRequerido = Val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
            TmpStr = Leer.GetValue("Hechizo" & Hechizo, "ManaRequeridoPerc")
            .ManaRequeridoPerc = Val(Replace(TmpStr, "%", ""))
            
            .RequireFullMana = Val(Leer.GetValue("Hechizo" & Hechizo, "RequiereManaCompleta"))
            
            'Barrin 30/9/03
            .StaRequerido = Val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
            
            .TargetUser = CBool(Val(Leer.GetValue("Hechizo" & Hechizo, "TargetUser")))
            .TargetNpc = CBool(Val(Leer.GetValue("Hechizo" & Hechizo, "TargetNpc")))
            .TargetObj = CBool(Val(Leer.GetValue("Hechizo" & Hechizo, "TargetObj")))
            .TargetTerrain = CBool(Val(Leer.GetValue("Hechizo" & Hechizo, "TargetTerrain")))

            
            frmCargando.Cargar.Value = frmCargando.Cargar.Value + 1
            
            .MagicCastPowerRequired = Val(Leer.GetValue("Hechizo" & Hechizo, "MagicCastPowerRequired"))
            
            .MinLevel = CByte(Val(Leer.GetValue("Hechizo" & Hechizo, "MinLevel")))
            
            
            .DamageOverTime.IsDot = CBool(Val(Leer.GetValue("Hechizo" & Hechizo, "IsDamageOverTime")))
            .DamageOverTime.TickCount = CInt(Val(Leer.GetValue("Hechizo" & Hechizo, "DotTickCount")))
            .DamageOverTime.TickInterval = CLng(Val(Leer.GetValue("Hechizo" & Hechizo, "DotTickInterval")))
            .DamageOverTime.WaitForFirstTick = CBool(Val(Leer.GetValue("Hechizo" & Hechizo, "DotWaitForFirstTick")))
            .DamageOverTime.MaxStackEffect = CInt(Val(Leer.GetValue("Hechizo" & Hechizo, "DotMaxStackEffect")))
            If (.DamageOverTime.MaxStackEffect <= 0) Then .DamageOverTime.MaxStackEffect = 1
            
            .IgnoreMagicDefensePerc = CByte(Val(Leer.GetValue("Hechizo" & Hechizo, "IgnoreMagicDefensePerc")))
                        
            .LifeLeechPerc = CByte(Val(Leer.GetValue("Hechizo" & Hechizo, "LifeLeechPerc")))
            .SpellCastInterval = Val(Leer.GetValue("Hechizo" & Hechizo, "SpellCastInterval"))
        End With
    Next Hechizo
    
    Set Leer = Nothing
    
    Exit Sub

ErrHandler:
    MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.Description
    End
 
End Sub

Sub LoadMotd()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim I As Integer
    
    MaxLines = Val(GetVar(ServerConfiguration.ResourcesPaths.Dats & "Motd.ini", "INIT", "NumLines"))
    
    ReDim MOTD(1 To MaxLines)
    For I = 1 To MaxLines
        MOTD(I).texto = GetVar(ServerConfiguration.ResourcesPaths.Dats & "Motd.ini", "Motd", "Line" & I)
        MOTD(I).Formato = vbNullString
    Next I

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadMotd de FileIO.bas")
End Sub

Public Sub DoBackUp()

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    On Error GoTo ErrHandler

    haciendoBK = True
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    
    Call LimpiarMundo
    Call WorldSave
    'Call modGuilds.CheckGuildsElections
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle(), IsUrgent:=True)
    
    haciendoBK = False
    
    'Log
    On Error Resume Next

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open ServerConfiguration.LogsPaths.GeneralPath & "BackUps.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time
    Close #nfile
  
    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DoBackUp de FileIO.bas")
End Sub

Public Sub GrabarMapa(ByVal Map As Long, ByRef MAPFILE As String)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2011
'10/08/2010 - Pato: Implemento el clsByteBuffer para el grabado de mapas
'28/10/2010:ZaMa - Ahora no se hace backup de los pretorianos.
'***************************************************
On Error GoTo ErrHandler
  

On Error Resume Next
    Dim FreeFileMap As Long
    Dim I As Long
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte
    Dim LoopC As Long
    Dim MapWriter As BinaryWriter
    Dim InfWriter As BinaryWriter
    Dim IniManager As clsIniManager
    Dim NpcInvalido As Boolean
    
    Set MapWriter = New BinaryWriter
    Set InfWriter = New BinaryWriter
    Set IniManager = New clsIniManager
    
    If FileExist(MAPFILE & ".map", vbNormal) Then
        Kill MAPFILE & ".map"
    End If
    
    If FileExist(MAPFILE & ".inf", vbNormal) Then
        Kill MAPFILE & ".inf"
    End If

    Dim Bytes() As Byte

    'map Header
    Call MapWriter.WriteInt16(MapInfo(Map).MapVersion)
    Call MapWriter.Write(StrConv(MiCabecera.Desc, vbFromUnicode), Len(MiCabecera.Desc))
    Call MapWriter.WriteInt32(MiCabecera.Crc)
    Call MapWriter.WriteInt32(MiCabecera.MagicWord)
    Call MapWriter.WriteReal64(0#)
    
    'inf Header
    Call InfWriter.WriteReal64(0#)
    Call InfWriter.WriteInt16(0)
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            With MapData(Map, X, Y)
                ByFlags = 0
                
                If .Blocked Then ByFlags = ByFlags Or 1
                If .Graphic(2) Then ByFlags = ByFlags Or 2
                If .Graphic(3) Then ByFlags = ByFlags Or 4
                If .Graphic(4) Then ByFlags = ByFlags Or 8
                If .Trigger Then ByFlags = ByFlags Or 16
                
                Call MapWriter.WriteInt8(ByFlags)
                
                Call MapWriter.WriteInt16(.Graphic(1))
                
                For LoopC = 2 To 4
                    If .Graphic(LoopC) Then _
                        Call MapWriter.WriteInt16(.Graphic(LoopC))
                Next LoopC
                
                If .Trigger Then _
                    Call MapWriter.WriteInt16(CInt(.Trigger))
                
                '.inf file
                ByFlags = 0
                
                If .ObjInfo.ObjIndex > 0 Then
                   If ObjData(.ObjInfo.ObjIndex).ObjType = eOBJType.otFogata Then
                        Call EraseObj(MAX_INVENTORY_OBJS, Map, X, Y)
                    End If
                End If
    
                If .TileExit.Map Then ByFlags = ByFlags Or 1
                
                ' No hacer backup de los NPCs inválidos (Pretorianos, Mascotas e Invocados)
                If .NpcIndex Then
                    NpcInvalido = (Npclist(.NpcIndex).NPCtype = eNPCType.Pretoriano) Or (Npclist(.NpcIndex).MaestroUser > 0)
                    
                    If Not NpcInvalido Then ByFlags = ByFlags Or 2
                End If
                
                If .ObjInfo.ObjIndex Then ByFlags = ByFlags Or 4
                
                Call InfWriter.WriteInt8(ByFlags)
                
                If .TileExit.Map Then
                    Call InfWriter.WriteInt16(.TileExit.Map)
                    Call InfWriter.WriteInt16(.TileExit.X)
                    Call InfWriter.WriteInt16(.TileExit.Y)
                End If
                
                If .NpcIndex And Not NpcInvalido Then _
                    Call InfWriter.WriteInt16(Npclist(.NpcIndex).Numero)
                
                If .ObjInfo.ObjIndex Then
                    Call InfWriter.WriteInt16(.ObjInfo.ObjIndex)
                    Call InfWriter.WriteInt16(.ObjInfo.Amount)
                End If
                
                NpcInvalido = False
            End With
        Next X
    Next Y
      
    'Open .map file
    FreeFileMap = FreeFile
     
    Open MAPFILE & ".Map" For Binary As FreeFileMap
        Call MapWriter.GetData(Bytes)
        Put FreeFileMap, , Bytes
    Close FreeFileMap
        
    Open MAPFILE & ".inf" For Binary As FreeFileMap
        Call InfWriter.GetData(Bytes)
        Put FreeFileMap, , Bytes
    Close FreeFileMap
    
    Set MapWriter = Nothing
    Set InfWriter = Nothing
 
    With MapInfo(Map)
        'write .dat file
        Call IniManager.ChangeValue("Mapa" & Map, "Name", .Name)
        
        Call IniManager.ChangeValue("Mapa" & Map, "NumMusic", .NumMusic)
        
        For I = 1 To .NumMusic
            Call IniManager.ChangeValue("Mapa" & Map, "MusicNum" & I, .Music(I))
        Next I
        
        Call IniManager.ChangeValue("Mapa" & Map, "MagiaSinefecto", .MagiaSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "InviSinEfecto", .InviSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "ResuSinEfecto", .ResuSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "StartPos", .StartPos.Map & "-" & .StartPos.X & "-" & .StartPos.Y)
        Call IniManager.ChangeValue("Mapa" & Map, "OnDeathGoTo", .OnDeathGoTo.Map & "-" & .OnDeathGoTo.X & "-" & .OnDeathGoTo.Y)
        Call IniManager.ChangeValue("Mapa" & Map, "MismoBando", .MismoBando)
        Call IniManager.ChangeValue("Mapa" & Map, "Reverb", .Reverb)
    
        Call IniManager.ChangeValue("Mapa" & Map, "Terreno", TerrainZoneByteToString(.Terreno))
        Call IniManager.ChangeValue("Mapa" & Map, "Zona", TerrainZoneByteToString(.Zona))
        Call IniManager.ChangeValue("Mapa" & Map, "Restringir", RestrictByteToString(.Restringir))
        Call IniManager.ChangeValue("Mapa" & Map, "BackUp", .BackUp)
    
        If .Pk Then
            Call IniManager.ChangeValue("Mapa" & Map, "Pk", "1")
        Else
            Call IniManager.ChangeValue("Mapa" & Map, "Pk", "0")
        End If
        
        Call IniManager.ChangeValue("Mapa" & Map, "OcultarSinEfecto", .OcultarSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "InvocarSinEfecto", .InvocarSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "InmovilizarSinEfecto", .InmovilizarSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "NoEncriptarMP", .NoEncriptarMP)
        Call IniManager.ChangeValue("Mapa" & Map, "RoboNpcsPermitido", .RoboNpcsPermitido)
        Call IniManager.ChangeValue("Mapa" & Map, "MapaTierra", .MapaTierra)
        
        Call IniManager.DumpFile(MAPFILE & ".dat")
    End With
    
    Set IniManager = Nothing
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GrabarMapa de FileIO.bas")
End Sub

Sub LoadBalance()
'***************************************************
'Author: Unknown
'Last Modification: 15/04/2010
'15/04/2010: ZaMa - Agrego recompensas faccionarias.
'22/07/2016: Anagrama - Agregadas constantes generales.
'***************************************************
On Error GoTo ErrHandler
  
    Dim TmpStr() As String
    Dim I As Long
    
    'Valores Generales
    With ConstantesBalance
        .LimiteNewbie = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "LimiteNewbie"))
        .MaxRep = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "MaxRep"))
        .MaxOro = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "MaxOro"))
        .MaxExp = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "MaxExp"))
        .MaxUsersMatados = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "MaxUsersMatados"))
        .MaxAtributos = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "MaxAtributos"))
        .MinAtributos = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "MinAtributos"))
        .MaxLvl = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "MaxLvl"))
        .MaxHp = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "MaxHP"))
        .MaxSta = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "MaxSta"))
        .MaxMan = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "MaxMan"))
        
        
        If Len(GetVar(DatPath & "Balance.dat", "GENERAL", "SelfWorkerMaps")) > 0 Then
            TmpStr = Split(GetVar(DatPath & "Balance.dat", "GENERAL", "SelfWorkerMaps"), ",")
            
            .SelfWorkerMapsQty = UBound(TmpStr) + 1
            ReDim .SelfWorkerMaps(1 To .SelfWorkerMapsQty)
            For I = 1 To UBound(TmpStr) + 1
                .SelfWorkerMaps(I) = CInt(TmpStr(I - 1))
            Next I
        Else
            .SelfWorkerMapsQty = 0
        End If
        
        .EluSkillInicial = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "EluSkillInicial"))
        .ExpAciertoSkill = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "ExpAciertoSkill"))
        .SkillExpCampfireSuccess = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "SkillExpCampfireSuccess"))
        .SkillExpNpcKilled = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "SkillExpNpcKilled"))
        
        .ExpFalloSkill = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "ExpFalloSkill"))
        .ModDefSegJerarquia = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "ModDefSegJerarquia"))
        .MinCrearPartyLevel = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "MinCrearPartyLevel"))
        .IntMoveAttack = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "IntMoveAttack"))
        
        .ModExpMultiplier = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "ModExpMultiplier"))
        .ModGoldMultiplier = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "ModGoldMultiplier"))
        .ModTrainingExpMultiplier = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "ModTrainingExpMultiplier"))
        
        .HomeWaitingTime = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "HomeWaitingTime"))
        .FactionMinLevel = CByte(Val(GetVar(DatPath & "Balance.dat", "GENERAL", "FactionMinLevel")))
        .FactionMaxRejoins = CByte(Val(GetVar(DatPath & "Balance.dat", "GENERAL", "FactionMaxRejoins")))
        
        ' Set the default values.
        If .ModExpMultiplier <= 0 Then .ModExpMultiplier = 1
        If .ModGoldMultiplier <= 0 Then .ModGoldMultiplier = 1
        If .ModTrainingExpMultiplier <= 0 Then .ModTrainingExpMultiplier = 1
        
        .MaxActiveTrapQty = Val(GetVar(DatPath & "Balance.dat", "GENERAL", "CantMaxTrapActivated"))
        
        .RankingMinLevel = Val(GetVar(DatPath & "Balance.dat", "RANKING", "RankingMinLevel"))
        .PlayerRankingStartingPoints = Val(GetVar(DatPath & "Balance.dat", "RANKING", "PlayerRankingStartingPoints"))
        .GuildRankingStartingPoints = Val(GetVar(DatPath & "Balance.dat", "RANKING", "GuildRankingStartingPoints"))
        .RankingSkewDistance = Val(GetVar(DatPath & "Balance.dat", "RANKING", "RankingSkewDistance"))
        
        ' Load the prohibited spells during duels.
        If Len(GetVar(DatPath & "Balance.dat", "GENERAL", "DuelProhibitedSpells")) > 0 Then
            TmpStr = Split(GetVar(DatPath & "Balance.dat", "GENERAL", "DuelProhibitedSpells"), ",")
            
            .DuelProhibitedSpellsQty = UBound(TmpStr) + 1
            
            ReDim .DuelProhibitedSpells(1 To .DuelProhibitedSpellsQty)
            For I = 1 To UBound(TmpStr) + 1
                .DuelProhibitedSpells(I) = CInt(TmpStr(I - 1))
            Next I
        Else
            .DuelProhibitedSpellsQty = 0
        End If
        
        ' Alignment matrix
        ' This matrix determines what kind of actions one alignment can execute against another. For example, can a Neutral player attack/rob a Royal character?
        
        ' Alignment attack matrix
        .AlignmentAttackActionMatrix(eCharacterAlignment.Newbie, eCharacterAlignment.Newbie) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "AttackActionMatrix" & eCharacterAlignment.Newbie & "_" & eCharacterAlignment.Newbie)))
        .AlignmentAttackActionMatrix(eCharacterAlignment.Newbie, eCharacterAlignment.Neutral) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "AttackActionMatrix" & eCharacterAlignment.Newbie & "_" & eCharacterAlignment.Neutral)))
        .AlignmentAttackActionMatrix(eCharacterAlignment.Newbie, eCharacterAlignment.FactionRoyal) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "AttackActionMatrix" & eCharacterAlignment.Newbie & "_" & eCharacterAlignment.FactionRoyal)))
        .AlignmentAttackActionMatrix(eCharacterAlignment.Newbie, eCharacterAlignment.FactionLegion) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "AttackActionMatrix" & eCharacterAlignment.Newbie & "_" & eCharacterAlignment.FactionLegion)))
        
        .AlignmentAttackActionMatrix(eCharacterAlignment.Neutral, eCharacterAlignment.Newbie) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "AttackActionMatrix" & eCharacterAlignment.Neutral & "_" & eCharacterAlignment.Newbie)))
        .AlignmentAttackActionMatrix(eCharacterAlignment.Neutral, eCharacterAlignment.Neutral) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "AttackActionMatrix" & eCharacterAlignment.Neutral & "_" & eCharacterAlignment.Neutral)))
        .AlignmentAttackActionMatrix(eCharacterAlignment.Neutral, eCharacterAlignment.FactionRoyal) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "AttackActionMatrix" & eCharacterAlignment.Neutral & "_" & eCharacterAlignment.FactionRoyal)))
        .AlignmentAttackActionMatrix(eCharacterAlignment.Neutral, eCharacterAlignment.FactionLegion) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "AttackActionMatrix" & eCharacterAlignment.Neutral & "_" & eCharacterAlignment.FactionLegion)))
        
        .AlignmentAttackActionMatrix(eCharacterAlignment.FactionRoyal, eCharacterAlignment.Newbie) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "AttackActionMatrix" & eCharacterAlignment.FactionRoyal & "_" & eCharacterAlignment.Newbie)))
        .AlignmentAttackActionMatrix(eCharacterAlignment.FactionRoyal, eCharacterAlignment.Neutral) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "AttackActionMatrix" & eCharacterAlignment.FactionRoyal & "_" & eCharacterAlignment.Neutral)))
        .AlignmentAttackActionMatrix(eCharacterAlignment.FactionRoyal, eCharacterAlignment.FactionRoyal) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "AttackActionMatrix" & eCharacterAlignment.FactionRoyal & "_" & eCharacterAlignment.FactionRoyal)))
        .AlignmentAttackActionMatrix(eCharacterAlignment.FactionRoyal, eCharacterAlignment.FactionLegion) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "AttackActionMatrix" & eCharacterAlignment.FactionRoyal & "_" & eCharacterAlignment.FactionLegion)))
        
        .AlignmentAttackActionMatrix(eCharacterAlignment.FactionLegion, eCharacterAlignment.Newbie) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "AttackActionMatrix" & eCharacterAlignment.FactionLegion & "_" & eCharacterAlignment.Newbie)))
        .AlignmentAttackActionMatrix(eCharacterAlignment.FactionLegion, eCharacterAlignment.Neutral) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "AttackActionMatrix" & eCharacterAlignment.FactionLegion & "_" & eCharacterAlignment.Neutral)))
        .AlignmentAttackActionMatrix(eCharacterAlignment.FactionLegion, eCharacterAlignment.FactionRoyal) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "AttackActionMatrix" & eCharacterAlignment.FactionLegion & "_" & eCharacterAlignment.FactionRoyal)))
        .AlignmentAttackActionMatrix(eCharacterAlignment.FactionLegion, eCharacterAlignment.FactionLegion) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "AttackActionMatrix" & eCharacterAlignment.FactionLegion & "_" & eCharacterAlignment.FactionLegion)))
        
        ' Alignment help matrix
        .AlignmentHelpActionMatrix(eCharacterAlignment.Newbie, eCharacterAlignment.Newbie) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "HelpActionMatrix" & eCharacterAlignment.Newbie & "_" & eCharacterAlignment.Newbie)))
        .AlignmentHelpActionMatrix(eCharacterAlignment.Newbie, eCharacterAlignment.Neutral) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "HelpActionMatrix" & eCharacterAlignment.Newbie & "_" & eCharacterAlignment.Neutral)))
        .AlignmentHelpActionMatrix(eCharacterAlignment.Newbie, eCharacterAlignment.FactionRoyal) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "HelpActionMatrix" & eCharacterAlignment.Newbie & "_" & eCharacterAlignment.FactionRoyal)))
        .AlignmentHelpActionMatrix(eCharacterAlignment.Newbie, eCharacterAlignment.FactionLegion) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "HelpActionMatrix" & eCharacterAlignment.Newbie & "_" & eCharacterAlignment.FactionLegion)))
        
        .AlignmentHelpActionMatrix(eCharacterAlignment.Neutral, eCharacterAlignment.Newbie) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "HelpActionMatrix" & eCharacterAlignment.Neutral & "_" & eCharacterAlignment.Newbie)))
        .AlignmentHelpActionMatrix(eCharacterAlignment.Neutral, eCharacterAlignment.Neutral) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "HelpActionMatrix" & eCharacterAlignment.Neutral & "_" & eCharacterAlignment.Neutral)))
        .AlignmentHelpActionMatrix(eCharacterAlignment.Neutral, eCharacterAlignment.FactionRoyal) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "HelpActionMatrix" & eCharacterAlignment.Neutral & "_" & eCharacterAlignment.FactionRoyal)))
        .AlignmentHelpActionMatrix(eCharacterAlignment.Neutral, eCharacterAlignment.FactionLegion) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "HelpActionMatrix" & eCharacterAlignment.Neutral & "_" & eCharacterAlignment.FactionLegion)))
        
        .AlignmentHelpActionMatrix(eCharacterAlignment.FactionRoyal, eCharacterAlignment.Newbie) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "HelpActionMatrix" & eCharacterAlignment.FactionRoyal & "_" & eCharacterAlignment.Newbie)))
        .AlignmentHelpActionMatrix(eCharacterAlignment.FactionRoyal, eCharacterAlignment.Neutral) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "HelpActionMatrix" & eCharacterAlignment.FactionRoyal & "_" & eCharacterAlignment.Neutral)))
        .AlignmentHelpActionMatrix(eCharacterAlignment.FactionRoyal, eCharacterAlignment.FactionRoyal) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "HelpActionMatrix" & eCharacterAlignment.FactionRoyal & "_" & eCharacterAlignment.FactionRoyal)))
        .AlignmentHelpActionMatrix(eCharacterAlignment.FactionRoyal, eCharacterAlignment.FactionLegion) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "HelpActionMatrix" & eCharacterAlignment.FactionRoyal & "_" & eCharacterAlignment.FactionLegion)))
        
        .AlignmentHelpActionMatrix(eCharacterAlignment.FactionLegion, eCharacterAlignment.Newbie) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "HelpActionMatrix" & eCharacterAlignment.FactionLegion & "_" & eCharacterAlignment.Newbie)))
        .AlignmentHelpActionMatrix(eCharacterAlignment.FactionLegion, eCharacterAlignment.Neutral) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "HelpActionMatrix" & eCharacterAlignment.FactionLegion & "_" & eCharacterAlignment.Neutral)))
        .AlignmentHelpActionMatrix(eCharacterAlignment.FactionLegion, eCharacterAlignment.FactionRoyal) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "HelpActionMatrix" & eCharacterAlignment.FactionLegion & "_" & eCharacterAlignment.FactionRoyal)))
        .AlignmentHelpActionMatrix(eCharacterAlignment.FactionLegion, eCharacterAlignment.FactionLegion) = CBool(Val(GetVar(DatPath & "Balance.dat", "ALIGNMENTS", "HelpActionMatrix" & eCharacterAlignment.FactionLegion & "_" & eCharacterAlignment.FactionLegion)))
   
    End With
    
    ReDim Preserve ConstantesMeditations(3, ConstantesBalance.MaxLvl) As Long
    Dim PreviousValue As Long, DefaultMeditation As Long
    DefaultMeditation = Val(GetVar(DatPath & "Balance.dat", "MEDITATIONS", "MeditationDefault"))
    For I = 1 To ConstantesBalance.MaxLvl
        ConstantesMeditations(eCharacterAlignment.Newbie, I) = Val(GetVar(DatPath & "Balance.dat", "MEDITATIONS", "Meditation_" & eCharacterAlignment.Newbie & "_" & I))
        If ConstantesMeditations(eCharacterAlignment.Newbie, I) = 0 And I > 1 Then
            ConstantesMeditations(eCharacterAlignment.Newbie, I) = ConstantesMeditations(eCharacterAlignment.Newbie, I - 1)
        End If
    
        ConstantesMeditations(eCharacterAlignment.Neutral, I) = Val(GetVar(DatPath & "Balance.dat", "MEDITATIONS", "Meditation_" & eCharacterAlignment.Neutral & "_" & I))
        If ConstantesMeditations(eCharacterAlignment.Neutral, I) = 0 And I > 1 Then
            ConstantesMeditations(eCharacterAlignment.Neutral, I) = ConstantesMeditations(eCharacterAlignment.Neutral, I - 1)
        End If
        
        ConstantesMeditations(eCharacterAlignment.FactionRoyal, I) = Val(GetVar(DatPath & "Balance.dat", "MEDITATIONS", "Meditation_" & eCharacterAlignment.FactionRoyal & "_" & I))
        If ConstantesMeditations(eCharacterAlignment.FactionRoyal, I) = 0 And I > 1 Then
            ConstantesMeditations(eCharacterAlignment.FactionRoyal, I) = ConstantesMeditations(eCharacterAlignment.FactionRoyal, I - 1)
        End If
        
        ConstantesMeditations(eCharacterAlignment.FactionLegion, I) = Val(GetVar(DatPath & "Balance.dat", "MEDITATIONS", "Meditation_" & eCharacterAlignment.FactionLegion & "_" & I))
        If ConstantesMeditations(eCharacterAlignment.FactionLegion, I) = 0 And I > 1 Then
            ConstantesMeditations(eCharacterAlignment.FactionLegion, I) = ConstantesMeditations(eCharacterAlignment.FactionLegion, I - 1)
        End If
    Next I
    
    'Modificadores de Raza
    For I = 1 To NUMRAZAS
        With ModRaza(I)
            .Fuerza = Val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(I) + "Fuerza"))
            .Agilidad = Val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(I) + "Agilidad"))
            .Inteligencia = Val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(I) + "Inteligencia"))
            .Carisma = Val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(I) + "Carisma"))
            .Constitucion = Val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(I) + "Constitucion"))
        End With
    Next I
        
    'Distribución de Vida
    For I = 1 To 5
        DistribucionEnteraVida(I) = Val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "E" + CStr(I)))
    Next I
    For I = 1 To 4
        DistribucionSemienteraVida(I) = Val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "S" + CStr(I)))
    Next I
    
    'Extra
    PorcentajeRecuperoMana = Val(GetVar(DatPath & "Balance.dat", "EXTRA", "PorcentajeRecuperoMana"))

    'Party
    ExponenteNivelParty = Val(GetVar(DatPath & "Balance.dat", "PARTY", "ExponenteNivelParty"))
    
    ' Recompensas faccionarias
    For I = 1 To NUM_RANGOS_FACCION
        RecompensaFacciones(I - 1) = Val(GetVar(DatPath & "Balance.dat", "RECOMPENSAFACCION", "Rango" & I))
    Next I
    
    ' Experiencia p/Nivel
    For I = 1 To 50
        TablaExperiencia(I) = CLng(Val(GetVar(DatPath & "Balance.dat", "EXPERIENCIA", "Nivel" & I)))
    Next I

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadBalance de FileIO.bas")
End Sub

Sub LoadResources()
On Error GoTo ErrHandler

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de recursos extraibles."
    
    '*****************************************************************
    'Carga la lista de recursos
    '*****************************************************************
    Dim Index As Integer
    Dim Leer As clsIniManager
    Set Leer = New clsIniManager
    
    Dim NumResources As Integer
    
    Call Leer.Initialize(DatPath & "Resources.dat")
    
    NumResources = Val(Leer.GetValue("INIT", "NumResources"))
    
    frmCargando.Cargar.Min = 0
    frmCargando.Cargar.Max = NumResources
    frmCargando.Cargar.Value = 0
    
    ReDim Preserve Resources(1 To NumResources) As tResource
    
    For Index = 1 To NumResources

        With Resources(Index)
            .ResourceNumber = Index
            .ObjIndex = Val(Leer.GetValue("Resource" & Index, "ObjIndex"))
            .ExtractionProbability = Val(Leer.GetValue("Resource" & Index, "ExtractionProbability"))
            .MinPerTickWorker = Val(Leer.GetValue("Resource" & Index, "MinPerTickWorker"))
            .MaxPerTickWorker = Val(Leer.GetValue("Resource" & Index, "MaxPerTickWorker"))
            .MinPerTickOther = Val(Leer.GetValue("Resource" & Index, "MinPerTickOther"))
            .MaxPerTickOther = Val(Leer.GetValue("Resource" & Index, "MaxPerTickOther"))
            .MaxAvailableQuantity = Val(Leer.GetValue("Resource" & Index, "MaxAvailableQuantity"))
            .MinToolPower = Val(Leer.GetValue("Resource" & Index, "MinToolPower"))
            
            If .MaxAvailableQuantity <= 0 Then .UnlimitedResource = True
        End With
        
        frmCargando.Cargar.Value = frmCargando.Cargar.Value + 1
    Next Index
    
    Set Leer = Nothing

    Exit Sub


ErrHandler:
    MsgBox "error cargando recurso " & Index & ": " & Err.Number & ": " & Err.Description

End Sub
Sub LoadProfessions()
'***************************************************
On Error GoTo ErrHandler

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de las profesiones."
    
    '*****************************************************************
    'Carga la lista de objetos
    '*****************************************************************
    Dim Index As Byte, J As Byte
    Dim Leer As clsIniManager
    Dim RecipeFilePath As String
    Dim NumRecipes As Integer
    Set Leer = New clsIniManager
    
    
    Dim NumProfessions As Integer
    
    Call Leer.Initialize(DatPath & "Professions.dat")
    
    NumProfessions = Val(Leer.GetValue("INIT", "NumProfessions"))
    
    frmCargando.Cargar.Min = 0
    frmCargando.Cargar.Max = NumProfessions
    frmCargando.Cargar.Value = 0
    
    
    ReDim Preserve Professions(1 To NumProfessions) As tProfession
    
    For Index = 1 To NumProfessions

        With Professions(Index)
            .Name = Leer.GetValue("Profession" & Index, "Name")
            .Enabled = Val(Leer.GetValue("Profession" & Index, "Enabled"))
            .SkillNumber = Val(Leer.GetValue("Profession" & Index, "SkillNumber"))
            .SkillExpSuccess = Val(Leer.GetValue("Profession" & Index, "SkillExpSuccess"))
            .SkillExpFailure = Val(Leer.GetValue("Profession" & Index, "SkillExpFailure"))
            .RequiredStaminaWorker = Val(Leer.GetValue("Profession" & Index, "RequiredStaminaWorker"))
            .RequiredStaminaOther = Val(Leer.GetValue("Profession" & Index, "RequiredStaminaOther"))
            .MinRemovableResourcesPercent = Val(Leer.GetValue("Profession" & Index, "MinRemovableResourcesPercent"))
            .MaxRemovableResourcesPercent = Val(Leer.GetValue("Profession" & Index, "MaxRemovableResourcesPercent"))
            .EnabledInSafeZone = Val(Leer.GetValue("Profession" & Index, "EnabledInSafeZone"))
            
            .SuccessFx = Val(Leer.GetValue("Profession" & Index, "SuccessFx"))
            
            ' Load recipe groups
            .CraftingRecipeGroupsQty = Val(Leer.GetValue("Profession" & Index, "RecipesGroups"))
            If .CraftingRecipeGroupsQty > 0 Then
                ReDim .CraftingRecipeGroups(1 To .CraftingRecipeGroupsQty)
                
                For J = 1 To .CraftingRecipeGroupsQty
                    .CraftingRecipeGroups(J).DatFileName = Leer.GetValue("Profession" & Index, "RecipesGroup" & J & "FileName")
                    .CraftingRecipeGroups(J).TabTitle = Leer.GetValue("Profession" & Index, "RecipesGroup" & J & "Tab")
                    .CraftingRecipeGroups(J).TabImage = Leer.GetValue("Profession" & Index, "RecipesGroup" & J & "TabImage")
                    
                    RecipeFilePath = ServerConfiguration.ResourcesPaths.Dats & .CraftingRecipeGroups(J).DatFileName
                    
                    ' Load the crafting recipe files
                    Call LoadProfessionCraftingFile(RecipeFilePath, Index, J)
                Next J
            End If
        End With
        
        frmCargando.Cargar.Value = frmCargando.Cargar.Value + 1
    Next Index
    
    
    Set Leer = Nothing

    Exit Sub


ErrHandler:
    Set Leer = Nothing
    
    MsgBox "Error cargando profesion " & Index & ": " & Err.Number & ": " & Err.Description
End Sub

Sub LoadProfessionCraftingFile(ByRef RecipeFilePath As String, ByRef Profession As Byte, ByRef ProfessionCraftingGroup As Byte)
    Dim IniRecipesReader As clsIniManager
    Dim Tmp As String
    Dim I As Integer, J As Integer
    Dim K As Integer
    
    With Professions(Profession).CraftingRecipeGroups(ProfessionCraftingGroup)
        
        If Not FileExist(RecipeFilePath, vbArchive) Then
            Call MsgBox("No se pudo encontrar el archivo " & RecipeFilePath & ", el servidor no puede funcionar sin él.")
            End
        End If
        
        Set IniRecipesReader = New clsIniManager
        Call IniRecipesReader.Initialize(RecipeFilePath)
    
        .RecipesQty = Val(IniRecipesReader.GetValue("INIT", "Recipes"))
        .ProfessionType = Val(IniRecipesReader.GetValue("INIT", "ProfessionType"))
        
        If .RecipesQty > 0 Then
            ReDim .Recipes(1 To .RecipesQty)
            For K = 1 To .RecipesQty
                .Recipes(K).ObjIndex = Val(IniRecipesReader.GetValue("OBJ" & K, "Item"))
                .Recipes(K).RecipeIndex = Val(IniRecipesReader.GetValue("OBJ" & K, "RecipeIndex"))
                .Recipes(K).CraftingProbability = Val(IniRecipesReader.GetValue("OBJ" & K, "CraftingProbability"))
                .Recipes(K).BlacksmithSkillNeeded = Val(IniRecipesReader.GetValue("OBJ" & K, "BlacksmithSkillNeeded"))
                .Recipes(K).CarpenterSkillNeeded = Val(IniRecipesReader.GetValue("OBJ" & K, "CarpenterSkillNeeded"))
                .Recipes(k).TailoringSkillNeeded = Val(IniRecipesReader.GetValue("OBJ" & k, "TailorSkillNeeded"))
                .Recipes(k).ProduceAmount = Val(IniRecipesReader.GetValue("OBJ" & k, "ProduceAmount"))
                
                If .Recipes(k).ProduceAmount <= 0 Then .Recipes(k).ProduceAmount = 1
                If .Recipes(k).ProduceAmount > MAX_INVENTORY_OBJS Then .Recipes(k).ProduceAmount = MAX_INVENTORY_OBJS
                                          
                .Recipes(K).MaterialsQty = Val(IniRecipesReader.GetValue("OBJ" & K, "MaterialsQty"))
                If .Recipes(K).MaterialsQty > 0 Then
                    ReDim .Recipes(K).Materials(1 To .Recipes(K).MaterialsQty)
                    For I = 1 To .Recipes(K).MaterialsQty
                        Tmp = IniRecipesReader.GetValue("OBJ" & K, "Material" & I)
                        .Recipes(K).Materials(I).ObjIndex = Val(ReadField(1, Tmp, Asc("-")))
                        .Recipes(K).Materials(I).Amount = Val(ReadField(2, Tmp, Asc("-")))
                    Next I
                Else
                    Erase .Recipes(K).Materials
                End If
            Next K
        End If
    End With
   
    
    Set IniRecipesReader = Nothing
End Sub


Sub LoadOBJData()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'¡¡¡¡ NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer desde el OBJ.DAT se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

On Error GoTo ErrHandler

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."
    
    '*****************************************************************
    'Carga la lista de objetos
    '*****************************************************************
    Dim I As Integer
    Dim N As Integer
    Dim S As String
    
    Dim prob As Long
    
    Dim Object As Integer
    Dim Leer As clsIniManager
    Set Leer = New clsIniManager
    
    Dim Counter As Long
    
    Call Leer.Initialize(DatPath & "Obj.dat")
    
    'obtiene el numero de obj
    NumObjDatas = Val(Leer.GetValue("INIT", "NumObjs"))
    
    frmCargando.Cargar.Min = 0
    frmCargando.Cargar.Max = NumObjDatas
    frmCargando.Cargar.Value = 0
    
    
    ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
    
    
    'Llena la lista
    For Object = 1 To NumObjDatas

        With ObjData(Object)
            .Name = Leer.GetValue("OBJ" & Object, "Name")
            
            'Pablo (ToxicWaste) Log de Objetos.
            .Log = Val(Leer.GetValue("OBJ" & Object, "Log"))
            .NoLog = Val(Leer.GetValue("OBJ" & Object, "NoLog"))
            '07/09/07
            
            .GrhIndex = Val(Leer.GetValue("OBJ" & Object, "GrhIndex"))
            If .GrhIndex = 0 Then
                .GrhIndex = .GrhIndex
            End If
            
            .ObjType = Val(Leer.GetValue("OBJ" & Object, "ObjType"))

            .Newbie = Val(Leer.GetValue("OBJ" & Object, "Newbie"))
            .Agarrable = Val(Leer.GetValue("OBJ" & Object, "Agarrable"))
            
            .Luminous = Val(Leer.GetValue("OBJ" & Object, "Luminous"))
            .ItemGM = Val(Leer.GetValue("OBJ" & Object, "ItemGM"))
            
            .CanBeTransparent = Val(Leer.GetValue("OBJ" & Object, "CanBeTransparent"))
            .ProfessionType = CLng(Val(Leer.GetValue("OBJ" & Object, "ProfessionType")))
            
            .MinimumLevel = Val(Leer.GetValue("OBJ" & Object, "MinimumLevel"))
            .Perforation = Val(Leer.GetValue("OBJ" & Object, "PerforationAmount"))
            
            If .Luminous Then
                .LightOffsetX = CInt(Val(Leer.GetValue("OBJ" & Object, "LightOffsetX")))
                .LightOffsetY = CInt(Val(Leer.GetValue("OBJ" & Object, "LightOffsetY")))
                .LightSize = CInt(Val(Leer.GetValue("OBJ" & Object, "LightSize")))
            End If
            
            Select Case .ObjType
                Case eOBJType.otArmadura
                    .Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
                    .LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                
                Case eOBJType.otESCUDO
                    .ShieldAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
                    
                Case eOBJType.otCASCO
                    .CascoAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
                
                Case eOBJType.otWeapon
                    .WeaponAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .Apuñala = Val(Leer.GetValue("OBJ" & Object, "Apuñala"))
                    .Envenena = Val(Leer.GetValue("OBJ" & Object, "Envenena"))
                    .MaxHit = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHit = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    .proyectil = Val(Leer.GetValue("OBJ" & Object, "Proyectil"))
                    .Municion = Val(Leer.GetValue("OBJ" & Object, "Municiones"))
                    .MagicCastPower = Val(Leer.GetValue("OBJ" & Object, "MagicCastPower"))
                    .MagicDamageBonus = Val(Leer.GetValue("OBJ" & Object, "MagicDamageBonus"))
                    .TwoHanded = Val(Leer.GetValue("OBJ" & Object, "TwoHanded"))
                    
                    .LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
                    
                    .WeaponRazaEnanaAnim = Val(Leer.GetValue("OBJ" & Object, "RazaEnanaAnim"))
                    .RequiredStamina = Val(Leer.GetValue("OBJ" & Object, "RequiredStamina"))
                    
                    .SplashDamage = CBool(Val(Leer.GetValue("OBJ" & Object, "SplashDamage")))
                    .SplashDamageType = Val(Leer.GetValue("OBJ" & Object, "SplashDamageType"))
                    .SplashDamageReduction = Val(Leer.GetValue("OBJ" & Object, "SplashDamageReduction"))
                
                Case eOBJType.otInstrumentos
                    .Snd1 = Val(Leer.GetValue("OBJ" & Object, "SND1"))
                    .Snd2 = Val(Leer.GetValue("OBJ" & Object, "SND2"))
                    .Snd3 = Val(Leer.GetValue("OBJ" & Object, "SND3"))

                    .Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
                
                Case eOBJType.otMinerales
                    .MinSkill = Val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                
                Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
                    .IndexAbierta = Val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
                    .IndexCerrada = Val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
                    .IndexCerradaLlave = Val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
                
                Case otPociones
                
                    .TipoPocion = Val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
                    .DuracionEfecto = Val(Leer.GetValue("OBJ" & Object, "EffectDuration"))
                    .MaxModificador = Val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
                    .MinModificador = Val(Leer.GetValue("OBJ" & Object, "MinModificador"))
                    
                    ' Adds mana?
                    Call Leer.GetMinMaxPercent("OBJ" & Object, "AffectsMana", .AffectsMana.Min, .AffectsMana.Max, .AffectsMana.IsPercent)
                    Call Leer.GetMinMaxPercent("OBJ" & Object, "AffectsHealth", .AffectsHealth.Min, .AffectsHealth.Max, .AffectsHealth.IsPercent)
                    Call Leer.GetMinMaxPercent("OBJ" & Object, "AffectsAgility", .AffectsAgility.Min, .AffectsAgility.Max, .AffectsAgility.IsPercent)
                    Call Leer.GetMinMaxPercent("OBJ" & Object, "AffectsStrength", .AffectsStrength.Min, .AffectsStrength.Max, .AffectsStrength.IsPercent)
                    
                    
                
                Case eOBJType.otBarcos
                    .MinSkill = Val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                    .MaxHit = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHit = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                
                Case eOBJType.otFlechas
                    .MaxHit = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHit = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    .Envenena = Val(Leer.GetValue("OBJ" & Object, "Envenena"))
                    .Paraliza = Val(Leer.GetValue("OBJ" & Object, "Paraliza"))
                    
                Case eOBJType.otAnillo
                    .LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .MaxHit = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHit = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    .MagicCastPower = Val(Leer.GetValue("OBJ" & Object, "MagicCastPower"))
                    .MagicDamageBonus = Val(Leer.GetValue("OBJ" & Object, "MagicDamageBonus"))
                    
                Case eOBJType.otTeleport
                    .Radio = Val(Leer.GetValue("OBJ" & Object, "Radio"))
                    
                Case eOBJType.otMochilas
                    .MochilaType = Val(Leer.GetValue("OBJ" & Object, "MochilaType"))
                    
                Case eOBJType.otForos
                    Call AddForum(Leer.GetValue("OBJ" & Object, "ID"))
                    
                ' Menues desplegables p/objeto
                Case eOBJType.otYunque
                    .MenuIndex = eMenues.ieYunque
                    
                Case eOBJType.otFragua
                    .MenuIndex = eMenues.ieFragua
                    
                ' Triggers
                Case eOBJType.otTrigger, eOBJType.otTrampa

                    With .Trigger
                        .Visible = CByte(Val(Leer.GetValue("OBJ" & Object, "TriggerVisible")))
                        .Dissapears = CByte(Val(Leer.GetValue("OBJ" & Object, "TriggerDissapears")))
                        .Animation = CInt(Val(Leer.GetValue("OBJ" & Object, "TriggerAnim")))
                        .CanDetect = CByte(Val(Leer.GetValue("OBJ" & Object, "TriggerDetect")))
                        .CanDisarm = CByte(Val(Leer.GetValue("OBJ" & Object, "TriggerDisarm")))
                        .CanTake = CByte(Val(Leer.GetValue("OBJ" & Object, "TriggerCanTake")))
                        .AffectNpc = CBool(Val(Leer.GetValue("OBJ" & Object, "TriggerAffectNpc")))
                        .AffectUser = CBool(Val(Leer.GetValue("OBJ" & Object, "TriggerAffectUser")))
                        
                        .ActivationMessage = Leer.GetValue("OBJ" & Object, "TriggerMessage")
                        
                        .NumSpells = CByte(Val(Leer.GetValue("OBJ" & Object, "TriggerNumSpells")))
                        
                        If .NumSpells > 0 Then
                            
                            ReDim .Spells(1 To .NumSpells)
                            
                            For Counter = 1 To .NumSpells
                                With .Spells(Counter)
                                    .Index = CInt(Val(Leer.GetValue("OBJ" & Object, "TriggerSpellIndex" & Counter)))
    
                                    .WAV = CInt(Val(Leer.GetValue("OBJ" & Object, "TriggerSpellWav" & Counter)))
                                    If .WAV = 0 Then .WAV = Hechizos(.Index).WAV
                                    
                                    .FXgrh = CInt(Val(Leer.GetValue("OBJ" & Object, "TriggerSpellFx" & Counter)))
                                    If .FXgrh = 0 Then .FXgrh = Hechizos(.Index).FXgrh
                                    
                                    .Loops = CByte(Val(Leer.GetValue("OBJ" & Object, "TriggerSpellLoops" & Counter)))
                                    If .Loops = 0 Then .Loops = Hechizos(.Index).Loops
                                    
                                    .Interval = CLng(Val(Leer.GetValue("OBJ" & Object, "TriggerSpellInterval" & Counter)))
                                    
                                    .MinHit = CInt(Val(Leer.GetValue("OBJ" & Object, "TriggerSpellMinHit" & Counter)))
                                    If .MinHit = 0 Then .MinHit = Hechizos(.Index).MinHp
                                    
                                    .MaxHit = CInt(Val(Leer.GetValue("OBJ" & Object, "TriggerSpellMaxHit" & Counter)))
                                    If .MaxHit = 0 Then .MaxHit = Hechizos(.Index).MaxHp
                                    
                                    .InvokeNpcIndex = CInt(Val(Leer.GetValue("OBJ" & Object, "TriggerSpellInvokeNpc" & Counter)))
                                    
                                    .DamageOverTime.IsDot = Hechizos(.Index).DamageOverTime.IsDot
                                    .DamageOverTime.TickCount = Hechizos(.Index).DamageOverTime.TickCount
                                    .DamageOverTime.TickInterval = Hechizos(.Index).DamageOverTime.TickInterval
                                    '. = CByte(Val(Leer.GetValue("OBJ" & Object, "TriggerSpellIsDot" & Counter)))
                                    
                                End With
                            Next Counter
                        End If
                    End With
                    
                    .TrapActivatedObject = CInt(Val(Leer.GetValue("OBJ" & Object, "ActivateObject")))
                    .TrapActivableLevelActivate = CInt(Val(Leer.GetValue("OBJ" & Object, "TrapActivableLevelActivate")))
                    .TrapActivableLevelDeactivate = CInt(Val(Leer.GetValue("OBJ" & Object, "TrapActivableLevelDeactivate")))
                    .TrapActivable = CByte(Val(Leer.GetValue("OBJ" & Object, "TrapActivable")))
                    
                    If .ObjType = eOBJType.otTrigger Then
                        ' Can't pick up
                        .Agarrable = 1
                    End If
                    
                Case eOBJType.otSurpriseBox
                    With .SurpriseDrops
                        .NroItems = CLng(Val(Leer.GetValue("OBJ" & Object, "NroItems")))
                        
                        If .NroItems > 0 Then
                            ReDim .Drop(1 To .NroItems) As tSurpriseObj
                            
                            prob = 0
                            
                            For I = 1 To .NroItems
                                Dim str As String
                                .Drop(I).ObjIndex = CInt(Val(ReadField(1, Leer.GetValue("OBJ" & Object, "Item" & I), Asc("-"))))
                                .Drop(I).Amount = CLng(Val(ReadField(2, Leer.GetValue("OBJ" & Object, "Item" & I), Asc("-"))))
                                
                                'prob = PROB_MULTIPLIER
                                'prob = prob + Val(ReadField(3, Leer.GetValue("OBJ" & Object, "Item" & I), Asc("-"))) * PROB_MULTIPLIER
                                
                                '.Drop(I).prob = Val(ReadField(3, Leer.GetValue("OBJ" & Object, "Item" & I), Asc("-"))) * PROB_MULTIPLIER
                            Next I
                        End If
                    End With
                Case eOBJType.otGuildBook
                    .Cupos = CLng(Val(Leer.GetValue("OBJ" & Object, "Cupos")))
                    
                ' obj properties for the resource gather system.
                Case eOBJType.otResource
                    .DepletedGrhIndex = CInt(Val(Leer.GetValue("OBJ" & Object, "DepletedGrhIndex")))
                    .MaxExtractedQuantity = CLng(Val(Leer.GetValue("OBJ" & Object, "MaxExtractedQuantity")))
                    .RespawnCooldown = CLng(Val(Leer.GetValue("OBJ" & Object, "RespawnCooldown")))
                    
                    .SoundNumber = CLng(Val(Leer.GetValue("OBJ" & Object, "SoundNumber")))
                    
                    'Default values
                    If .MaxExtractedQuantity = 0 Then
                        .MaxExtractedQuantity = 10000
                    End If
                    
                    '30 seconds
                    If .RespawnCooldown = 0 Then
                        .RespawnCooldown = 30
                    End If
                    
                    .NumResources = CLng(Val(Leer.GetValue("OBJ" & Object, "NumResources")))
                    If .NumResources > 0 Then
                        ReDim .Resources(1 To .NumResources) As tResource
                        Dim ResourceIndex As Integer
                        For I = 1 To .NumResources
                            ResourceIndex = Val(Leer.GetValue("OBJ" & Object, "Resource" & I))
                            .Resources(I) = Resources(ResourceIndex)
                        Next I
                    End If
                    
                    
                Case eOBJType.otTool
                    .WeaponAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .MaxHit = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHit = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    .ProfessionType = CByte(Val(Leer.GetValue("OBJ" & Object, "ProfessionType")))
                    .ToolPower = CByte(Val(Leer.GetValue("OBJ" & Object, "ToolPower")))
                    .SoundNumber = CLng(Val(Leer.GetValue("OBJ" & Object, "SoundNumber")))
                    .RequiredStamina = Leer.GetValue("OBJ" & Object, "RequiredStamina")
                    
                Case eOBJType.otFogata
                    .MaxStaRecoveryPerc = CByte(Val(Leer.GetValue("OBJ" & Object, "MaxStaRecoveryPerc")))
                    .DisappearTimeInSec = CByte(Val(Leer.GetValue("OBJ" & Object, "DisappearTimeInSec")))
                    .AllowResting = CBool(Val(Leer.GetValue("OBJ" & Object, "AllowResting")))
            End Select
            
            ' Menues desplegables p/objeto
            If Object = ConstantesItems.Leña Or Object = ConstantesItems.LeñaElfica Then
                .MenuIndex = eMenues.ieLenia
            ElseIf Object = ConstantesItems.Fogata Then
                .MenuIndex = eMenues.ieFogata
            ElseIf (Object = ConstantesItems.FogataApagada Or Object = ConstantesItems.RamitaElfica) Then
                .MenuIndex = eMenues.ieRamas
            End If
            
            .Ropaje = Val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
            
            .NumBodyNeutral = Val(Leer.GetValue("OBJ" & Object, "NumBodyNeutral"))
            .NumBodyRoyal = Val(Leer.GetValue("OBJ" & Object, "NumBodyRoyal"))
            .NumBodyLegion = Val(Leer.GetValue("OBJ" & Object, "NumBodyLegion"))
            
            
            .NumRopajeGenerico = Val(Leer.GetValue("OBJ" & Object, "NumRopajeGenerico"))
            .NumRopajeMujerAlto = Val(Leer.GetValue("OBJ" & Object, "NumRopajeMujerAlto"))
            .NumRopajeHombreAlto = Val(Leer.GetValue("OBJ" & Object, "NumRopajeHombreAlto"))
            .NumRopajeMujerBajo = Val(Leer.GetValue("OBJ" & Object, "NumRopajeMujerBajo"))
            .NumRopajeHombreBajo = Val(Leer.GetValue("OBJ" & Object, "NumRopajeHombreBajo"))
            .NumRopajeMujerDrow = Val(Leer.GetValue("OBJ" & Object, "NumRopajeMujerDrow"))
            .NumRopajeHombreDrow = Val(Leer.GetValue("OBJ" & Object, "NumRopajeHombreDrow"))
            
            
            .HechizoIndex = Val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
            
            .LingoteIndex = Val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
            
            .MineralIndex = Val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
            
            .MaxHp = Val(Leer.GetValue("OBJ" & Object, "MaxHP"))
            .MinHp = Val(Leer.GetValue("OBJ" & Object, "MinHP"))
            
            .Mujer = Val(Leer.GetValue("OBJ" & Object, "Mujer"))
            .Hombre = Val(Leer.GetValue("OBJ" & Object, "Hombre"))
            
            .MinHam = Val(Leer.GetValue("OBJ" & Object, "MinHam"))
            .MinSed = Val(Leer.GetValue("OBJ" & Object, "MinAgu"))
            
            .MinDef = Val(Leer.GetValue("OBJ" & Object, "MINDEF"))
            .MaxDef = Val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
            .Def = (.MinDef + .MaxDef) / 2
            
            .StabDamageReduction = Val(Leer.GetValue("OBJ" & Object, "StabDamageReduction"))
            .Valor = Val(Leer.GetValue("OBJ" & Object, "Valor"))
            .SalePrice = Val(Leer.GetValue("OBJ" & Object, "SalePrice"))
            
            .Crucial = Val(Leer.GetValue("OBJ" & Object, "Crucial"))
            
            .Cerrada = Val(Leer.GetValue("OBJ" & Object, "abierta"))
            If .Cerrada = 1 Then
                .Llave = Val(Leer.GetValue("OBJ" & Object, "Llave"))
                .clave = Val(Leer.GetValue("OBJ" & Object, "Clave"))
            End If
            
            'Puertas y llaves
            .clave = Val(Leer.GetValue("OBJ" & Object, "Clave"))
            
            .texto = Leer.GetValue("OBJ" & Object, "Texto")
            .GrhSecundario = Val(Leer.GetValue("OBJ" & Object, "VGrande"))
            
            .ForoID = Leer.GetValue("OBJ" & Object, "ID")
            
            .Acuchilla = Val(Leer.GetValue("OBJ" & Object, "Acuchilla"))
            .Critical = Val(Leer.GetValue("OBJ" & Object, "Critical"))
            
            .Guante = Val(Leer.GetValue("OBJ" & Object, "Guante"))
            
            .ItemQuest = Val(Leer.GetValue("OBJ" & Object, "ItemQuest"))
            
            'CHECK: !!! Esto es provisorio hasta que los de Dateo cambien los valores de string a numerico
            For I = 1 To NUMCLASES
                S = UCase$(Leer.GetValue("OBJ" & Object, "CP" & I))
                N = 1
                Do While LenB(S) > 0 And UCase$(ListaClases(N)) <> S
                    N = N + 1
                Loop
                .ClaseProhibida(I) = IIf(LenB(S) > 0, N, 0)
            Next I
            
            
            For I = 1 To NUMRAZAS
                S = UCase$(Leer.GetValue("OBJ" & Object, "RP" & I))
                N = 1
                 Do While LenB(S) > 0 And UCase$(ListaRazas(N)) <> S
                    N = N + 1
                Loop
                .RazaProhibida(I) = IIf(LenB(S) > 0, N, 0)
            Next I
            
            .DefensaMagicaMax = Val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
            .DefensaMagicaMin = Val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))
            
            .SkCarpinteria = Val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
            
            If .SkCarpinteria > 0 Then
                .Madera = Val(Leer.GetValue("OBJ" & Object, "Madera"))
                .MaderaElfica = Val(Leer.GetValue("OBJ" & Object, "MaderaElfica"))
            End If
            
            'Bebidas
            .MinSta = Val(Leer.GetValue("OBJ" & Object, "MinST"))
            
            .NoSeCae = Val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
            .NoSeTira = Val(Leer.GetValue("OBJ" & Object, "NoSeTira"))
            .NoRobable = Val(Leer.GetValue("OBJ" & Object, "NoRobable"))
            .NoComerciable = Val(Leer.GetValue("OBJ" & Object, "NoComerciable"))
            .Intransferible = Val(Leer.GetValue("OBJ" & Object, "Intransferible"))
            
            .ImpideParalizar = CByte(Val(Leer.GetValue("OBJ" & Object, "ImpideParalizar")))
            .ImpideInmobilizar = CByte(Val(Leer.GetValue("OBJ" & Object, "ImpideInmobilizar")))
            .ImpideAturdir = CByte(Val(Leer.GetValue("OBJ" & Object, "ImpideAturdir")))
            .ImpideCegar = CByte(Val(Leer.GetValue("OBJ" & Object, "ImpideCegar")))

            .Upgrade = Val(Leer.GetValue("OBJ" & Object, "Upgrade"))
            
            .CampfireObj = Val(Leer.GetValue("OBJ" & Object, "CampfireObj"))
            .MaxDistanceFromTarget = CByte(Val(Leer.GetValue("OBJ" & Object, "MaxDistanceFromTarget")))
            
            .SizeWidth = CByte(Val(Leer.GetValue("OBJ" & Object, "SizeWidth")))
            .SizeHeight = CByte(Val(Leer.GetValue("OBJ" & Object, "SizeHeight")))
            
            If .SizeWidth = 0 Then .SizeWidth = ModAreas.DEFAULT_ENTITY_WIDTH
            If .SizeHeight = 0 Then .SizeHeight = ModAreas.DEFAULT_ENTITY_HEIGHT
            
            frmCargando.Cargar.Value = frmCargando.Cargar.Value + 1
        End With
    Next Object
    
    
    Set Leer = Nothing
    
    ' Inicializo los foros faccionarios
    Call AddForum(FORO_CAOS_ID)
    Call AddForum(FORO_REAL_ID)
    
    Exit Sub

ErrHandler:
    MsgBox "error cargando objeto " & Object & ": " & Err.Number & ": " & Err.Description

End Sub

Sub LoadIntervals()
On Error GoTo ErrHandler

    Dim Leer As clsIniManager
    Set Leer = New clsIniManager
        
    Call Leer.Initialize(DatPath & "Intervals.dat")
    
    'Intervalos
    ServerConfiguration.Intervals.SanaIntervaloSinDescansar = Val(Leer.GetValue("INTERVALS", "SanaIntervaloSinDescansar"))
    FrmInterv.txtSanaIntervaloSinDescansar.Text = ServerConfiguration.Intervals.SanaIntervaloSinDescansar
    
    ServerConfiguration.Intervals.StaminaIntervaloSinDescansar = Val(Leer.GetValue("INTERVALS", "StaminaIntervaloSinDescansar"))
    FrmInterv.txtStaminaIntervaloSinDescansar.Text = ServerConfiguration.Intervals.StaminaIntervaloSinDescansar
    
    ServerConfiguration.Intervals.SanaIntervaloDescansar = Val(Leer.GetValue("INTERVALS", "SanaIntervaloDescansar"))
    FrmInterv.txtSanaIntervaloDescansar.Text = ServerConfiguration.Intervals.SanaIntervaloDescansar
    
    ServerConfiguration.Intervals.StaminaIntervaloDescansar = Val(Leer.GetValue("INTERVALS", "StaminaIntervaloDescansar"))
    FrmInterv.txtStaminaIntervaloDescansar.Text = ServerConfiguration.Intervals.StaminaIntervaloDescansar
    
    ServerConfiguration.Intervals.IntervaloSed = Val(Leer.GetValue("INTERVALS", "IntervaloSed"))
    FrmInterv.txtIntervaloSed.Text = ServerConfiguration.Intervals.IntervaloSed
    
    ServerConfiguration.Intervals.IntervaloHambre = Val(Leer.GetValue("INTERVALS", "IntervaloHambre"))
    FrmInterv.txtIntervaloHambre.Text = ServerConfiguration.Intervals.IntervaloHambre
    
    ServerConfiguration.Intervals.IntervaloVeneno = Val(Leer.GetValue("INTERVALS", "IntervaloVeneno"))
    FrmInterv.txtIntervaloVeneno.Text = ServerConfiguration.Intervals.IntervaloVeneno
    
    ServerConfiguration.Intervals.IntervaloPutrefaccionDmg = Val(Val(Leer.GetValue("INTERVALS", "IntervaloPutrefaccionDmg")))
    
    ServerConfiguration.Intervals.IntervaloParalizado = Val(Val(Leer.GetValue("INTERVALS", "IntervaloParalizado")))
    FrmInterv.txtIntervaloParalizado.Text = ServerConfiguration.Intervals.IntervaloParalizado
    
    ServerConfiguration.Intervals.IntervaloParalizadoReducido = Val(Val(Leer.GetValue("INTERVALS", "IntervaloParalizadoReducido")))
    
    ServerConfiguration.Intervals.IntervaloNPCParalizado = Val(Val(Leer.GetValue("INTERVALS", "IntervaloNPCParalizado")))
    
    ServerConfiguration.Intervals.IntervaloInvisible = Val(Leer.GetValue("INTERVALS", "IntervaloInvisible"))
    FrmInterv.txtIntervaloInvisible.Text = ServerConfiguration.Intervals.IntervaloInvisible
    
    ServerConfiguration.Intervals.IntervaloMimetismo = Val(Leer.GetValue("INTERVALS", "IntervaloMimetismo"))
    
    ServerConfiguration.Intervals.IntervaloFrio = Val(Leer.GetValue("INTERVALS", "IntervaloFrio"))
    FrmInterv.txtIntervaloFrio.Text = ServerConfiguration.Intervals.IntervaloFrio
    
    ServerConfiguration.Intervals.IntervaloLava = Val(Leer.GetValue("INTERVALS", "IntervaloLava"))
    
    ServerConfiguration.Intervals.IntervaloInvocacion = Val(Leer.GetValue("INTERVALS", "IntervaloInvocacion"))
    FrmInterv.txtInvocacion.Text = ServerConfiguration.Intervals.IntervaloInvocacion
    
    ServerConfiguration.Intervals.IntervaloIdleKick = modIntervals.FromMinutes((Leer.GetValue("INTERVALS", "IntervaloIdleKick")))
    FrmInterv.txtIntervaloIdleKick.Text = ServerConfiguration.Intervals.IntervaloIdleKick
    
    ServerConfiguration.Intervals.IntervaloOcultar = Val(Leer.GetValue("INTERVALS", "IntervaloOcultar"))
    FrmInterv.txtIntervaloOcultar.Text = ServerConfiguration.Intervals.IntervaloOcultar
    
    ServerConfiguration.Intervals.IntervaloInmunidad = Val(Leer.GetValue("INTERVALS", "IntervaloInmunidad"))
    FrmInterv.txtIntervaloInmunidad.Text = ServerConfiguration.Intervals.IntervaloInmunidad
    
    '&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    
    ServerConfiguration.Intervals.IntervaloPuedeSerAtacado = Val(Leer.GetValue("INTERVALS", "IntervaloPuedeSerAtacado"))
    ServerConfiguration.Intervals.IntervaloAtacable = Val(Leer.GetValue("INTERVALS", "IntervaloAtacable"))
    ServerConfiguration.Intervals.IntervaloOwnedNpc = Val(Leer.GetValue("INTERVALS", "IntervaloOwnedNpc"))
    
    ServerConfiguration.Intervals.IntervaloUserPuedeCastear = Val(Leer.GetValue("INTERVALS", "IntervaloLanzaHechizo"))
    FrmInterv.txtIntervaloLanzaHechizo.Text = ServerConfiguration.Intervals.IntervaloUserPuedeCastear
    
    frmMain.TIMER_AI.Interval = Val(Leer.GetValue("INTERVALS", "IntervaloNpcAI"))
    FrmInterv.txtAI.Text = frmMain.TIMER_AI.Interval
    
    frmMain.npcataca.Interval = Val(Leer.GetValue("INTERVALS", "IntervaloNpcPuedeAtacar"))
    FrmInterv.txtNPCPuedeAtacar.Text = frmMain.npcataca.Interval
    
    ServerConfiguration.Intervals.IntervaloUserPuedeTrabajar = Val(Leer.GetValue("INTERVALS", "IntervaloTrabajo"))
    FrmInterv.txtTrabajo.Text = ServerConfiguration.Intervals.IntervaloUserPuedeTrabajar
    
    ServerConfiguration.Intervals.IntervaloUserPuedeAtacar = Val(Leer.GetValue("INTERVALS", "IntervaloUserPuedeAtacar"))
    FrmInterv.txtPuedeAtacar.Text = ServerConfiguration.Intervals.IntervaloUserPuedeAtacar
    
    'TODO : Agregar estos intervalos al form!!!
    ServerConfiguration.Intervals.IntervaloMagiaGolpe = Val(Leer.GetValue("INTERVALS", "IntervaloMagiaGolpe"))
    ServerConfiguration.Intervals.IntervaloGolpeMagia = Val(Leer.GetValue("INTERVALS", "IntervaloGolpeMagia"))
    ServerConfiguration.Intervals.IntervaloGolpeUsar = Val(Leer.GetValue("INTERVALS", "IntervaloGolpeUsar"))
    
    frmMain.tLluvia.Interval = Val(Leer.GetValue("INTERVALS", "IntervaloPerdidaStaminaLluvia"))
    FrmInterv.txtIntervaloPerdidaStaminaLluvia.Text = frmMain.tLluvia.Interval
    
    MinutosWs = Val(Leer.GetValue("INTERVALS", "IntervaloWS"))
    If MinutosWs < 60 Then MinutosWs = 180
    
    MinutosMotd = Val(Leer.GetValue("INTERVALS", "MinutosMotd"))
    If MinutosMotd < 20 Then MinutosMotd = 20

    MinutosGuardarUsuarios = (Val(Leer.GetValue("INTERVALS", "IntervaloGuardarUsuarios")))
    
    ServerConfiguration.Intervals.IntervaloCerrarConexion = Val(Leer.GetValue("INTERVALS", "IntervaloCerrarConexion"))
    ServerConfiguration.Intervals.IntervaloUserPuedeUsar = Val(Leer.GetValue("INTERVALS", "IntervaloUserPuedeUsar"))
    ServerConfiguration.Intervals.IntervaloUserPuedeUsarU = Val(Leer.GetValue("INTERVALS", "IntervaloUserPuedeUsarU"))
    
    ServerConfiguration.Intervals.IntervaloFlechasCazadores = Val(Leer.GetValue("INTERVALS", "IntervaloFlechasCazadores"))
    
    ServerConfiguration.Intervals.IntervaloOculto = Val(Leer.GetValue("INTERVALS", "IntervaloOculto"))
    
    ServerConfiguration.Intervals.IntervalRequestPosition = Val(Leer.GetValue("INTERVALS", "IntervalRequestPosition"))
    ServerConfiguration.Intervals.IntervalMeditate = Val(Leer.GetValue("INTERVALS", "IntervalMeditate"))
    ServerConfiguration.Intervals.IntervalAction = Val(Leer.GetValue("INTERVALS", "IntervalAction"))
    ServerConfiguration.Intervals.IntervalWorkMacro = Val(Leer.GetValue("INTERVALS", "IntervalWorkMacro"))
    ServerConfiguration.Intervals.IntervalSpellMacro = Val(Leer.GetValue("INTERVALS", "IntervalSpellMacro"))
    
    '&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    
    
    Exit Sub

ErrHandler:
    MsgBox "Error cargando intervalos. " & Err.Number & ": " & Err.Description
    
End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim sSpaces As String ' This will hold the input that the program will retrieve
    Dim szReturn As String ' This will be the defaul value if the string is not found
      
    szReturn = vbNullString
      
    sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
      
      
    GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, File
      
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
  
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetVar de FileIO.bas")
End Function

Sub CargarBackUp()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup."
    
    Dim Map As Integer
    Dim tFileName As String
    
    On Error GoTo man
        
        NumMaps = Val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
        Call ModAreas.Initialise(NumMaps)
        
        frmCargando.Cargar.Min = 0
        frmCargando.Cargar.Max = NumMaps
        frmCargando.Cargar.Value = 0
        
        MapPath = ServerConfiguration.ResourcesPaths.Maps
        
        ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
        ReDim MapInfo(1 To NumMaps) As MapInfo
        
        For Map = 1 To NumMaps
            If Val(GetVar(MapPath & "Mapa" & Map & ".Dat", "Mapa" & Map, "BackUp")) <> 0 Then
                tFileName = ServerConfiguration.ResourcesPaths.WorldBackup & "\Mapa" & Map
                
                If Not FileExist(tFileName & ".*") Then 'Miramos que exista al menos uno de los 3 archivos, sino lo cargamos de la carpeta de los mapas
                    tFileName = MapPath & "Mapa" & Map
                End If
            Else
                tFileName = MapPath & "Mapa" & Map
            End If
                  
            Call CargarMapa(Map, tFileName)
            
            frmCargando.Cargar.Value = frmCargando.Cargar.Value + 1
            DoEvents
        Next Map
    
    Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)
 
End Sub

Sub LoadMapData()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas..."
    
    Dim Map As Integer
    Dim tFileName As String
    
    On Error GoTo man
        
        NumMaps = Val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
        Call ModAreas.Initialise(NumMaps)
        
        frmCargando.Cargar.Min = 0
        frmCargando.Cargar.Max = NumMaps
        frmCargando.Cargar.Value = 0
        
        MapPath = ServerConfiguration.ResourcesPaths.Maps
        
        
        ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
        ReDim MapInfo(1 To NumMaps) As MapInfo
          
        For Map = 1 To NumMaps
            
            tFileName = MapPath & "Mapa" & Map
            Call CargarMapa(Map, tFileName)
            
            frmCargando.Cargar.Value = frmCargando.Cargar.Value + 1
            DoEvents
        Next Map
    
    Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)

End Sub

Public Sub CargarMapa(ByVal Map As Long, ByRef MAPFl As String)
'***************************************************
'Author: Unknown
'Last Modification: 10/08/2010
'10/08/2010 - Pato: Implemento el clsByteBuffer y el clsIniManager para la carga de mapa
'***************************************************

On Error GoTo errh
    Dim hFile As Integer

    Dim X As Long
    Dim Y As Long

    Dim npcfile As String
    Dim Leer As clsIniManager
    Dim I As Long
    Dim MapReader As BinaryReader
    Dim InfReader As BinaryReader
    Dim Buff1() As Byte
    Dim Buff2() As Byte

    Dim PosType As Byte 'Si la posicion es de tierra (0) o agua (1).
    ReDim MapInfo(Map).NpcSpawnPos(0).Pos(0)
    ReDim MapInfo(Map).NpcSpawnPos(1).Pos(0)
    
    Dim NpcIndexTest As Integer
    
    hFile = FreeFile

    Open MAPFl & ".map" For Binary As #hFile
        Seek hFile, 1

        ReDim Buff1(LOF(hFile) - 1) As Byte
    
        Get #hFile, , Buff1
    Close hFile
    
    Open MAPFl & ".inf" For Binary As #hFile
        Seek hFile, 1

        ReDim Buff2(LOF(hFile) - 1) As Byte
    
        Get #hFile, , Buff2
    Close hFile
      
    'map Header
    Set MapReader = New BinaryReader
    Call MapReader.SetData(Buff1)
    
    Set InfReader = New BinaryReader
    Call InfReader.SetData(Buff2)
    
    MapInfo(Map).MapVersion = MapReader.ReadInt16
    
    Dim Header(0 To 254) As Byte
    Call MapReader.Read(Header(0), 255)

    MiCabecera.Desc = StrConv(Header, vbUnicode)
    MiCabecera.Crc = MapReader.ReadInt32()
    MiCabecera.MagicWord = MapReader.ReadInt32()

    Call MapReader.Skip(8) ' Double

    'inf Header
    Call InfReader.Skip(10) ' Double + Int

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            Call LoadMapDataForXAndY(Map, X, Y, MapReader, InfReader)
        Next X
    Next Y

    Set Leer = New clsIniManager
    Call Leer.Initialize(MAPFl & ".dat")

    With MapInfo(Map)
        .Name = Leer.GetValue("Mapa" & Map, "Name")
        
        If Val(Leer.GetValue("Mapa" & Map, "NumMusic")) > 0 Then
            .NumMusic = Leer.GetValue("Mapa" & Map, "NumMusic")
        
            ReDim .Music(1 To .NumMusic) As Long
        
            For I = 1 To .NumMusic
                .Music(I) = Val(Leer.GetValue("Mapa" & Map, "MusicNum" & I))
            Next I
        End If

        .StartPos.Map = Val(ReadField(1, Leer.GetValue("Mapa" & Map, "StartPos"), Asc("-")))
        .StartPos.X = Val(ReadField(2, Leer.GetValue("Mapa" & Map, "StartPos"), Asc("-")))
        .StartPos.Y = Val(ReadField(3, Leer.GetValue("Mapa" & Map, "StartPos"), Asc("-")))
        
        .OnDeathGoTo.Map = Val(ReadField(1, Leer.GetValue("Mapa" & Map, "OnDeathGoTo"), Asc("-")))
        .OnDeathGoTo.X = Val(ReadField(2, Leer.GetValue("Mapa" & Map, "OnDeathGoTo"), Asc("-")))
        .OnDeathGoTo.Y = Val(ReadField(3, Leer.GetValue("Mapa" & Map, "OnDeathGoTo"), Asc("-")))
        
        .MagiaSinEfecto = Val(Leer.GetValue("Mapa" & Map, "MagiaSinEfecto"))
        .InviSinEfecto = Val(Leer.GetValue("Mapa" & Map, "InviSinEfecto"))
        .ResuSinEfecto = Val(Leer.GetValue("Mapa" & Map, "ResuSinEfecto"))
        .OcultarSinEfecto = Val(Leer.GetValue("Mapa" & Map, "OcultarSinEfecto"))
        .InvocarSinEfecto = Val(Leer.GetValue("Mapa" & Map, "InvocarSinEfecto"))
        .InmovilizarSinEfecto = Val(Leer.GetValue("Mapa" & Map, "InmovilizarSinEfecto"))
        .MismoBando = Val(Leer.GetValue("Mapa" & Map, "MismoBando"))
        .Reverb = Val(Leer.GetValue("Mapa" & Map, "Reverb"))

        .NoEncriptarMP = Val(Leer.GetValue("Mapa" & Map, "NoEncriptarMP"))

        .RoboNpcsPermitido = Val(Leer.GetValue("Mapa" & Map, "RoboNpcsPermitido"))
        .MapaTierra = Val(Leer.GetValue("Mapa" & Map, "MapaTierra"))
        
        If Val(Leer.GetValue("Mapa" & Map, "Pk")) = 1 Then
            .Pk = True
        Else
            .Pk = False
        End If
        
        .NakedLosesEnergy = Leer.GetBooleanOrDefault("Mapa" & Map, "NakedLosesEnergy", True)
        .NakedLosesHealth = Leer.GetBooleanOrDefault("Mapa" & Map, "NakedLosesHealth", False)
        
        .Terreno = TerrainZoneStringToByte(Leer.GetValue("Mapa" & Map, "Terreno"))
        .Zona = TerrainZoneStringToByte(Leer.GetValue("Mapa" & Map, "Zona"))
        .Restringir = RestrictStringToByte(Leer.GetValue("Mapa" & Map, "Restringir"))
        .BackUp = Val(Leer.GetValue("Mapa" & Map, "BACKUP"))
        
        ' Can the user open a crafting store/self-worker store in this map?
        .CraftingStoreAllowed = False
        For I = 1 To ConstantesBalance.SelfWorkerMapsQty
            If Map = ConstantesBalance.SelfWorkerMaps(I) Then
                .CraftingStoreAllowed = True
                Exit For
            End If
        Next I
    End With
    
    ' If there are any Extractable Resource in the map, the we generate the empty slots needed for respawn
    Call GenerateEmptyResources(Map)
Exit Sub

errh:
    Call LogError("Error cargando mapa: " & Map & " - Pos: " & X & "," & Y & "." & Err.Description & " - NpcIndexTest: " & NpcIndexTest)

End Sub

Private Sub LoadMapDataForXAndY(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByRef MapReader As BinaryReader, ByRef InfReader As BinaryReader)

On Error GoTo ErrHandler

    Dim ByFlags As Byte
    Dim PosType  As Byte
    Dim NpcIndexTest As Integer
    
    With MapData(Map, X, Y)
        '.map file
        ByFlags = MapReader.ReadInt8

        If ByFlags And 1 Then .Blocked = 1

        .Graphic(1) = MapReader.ReadInt16

        'Layer 2 used?
        If ByFlags And 2 Then .Graphic(2) = MapReader.ReadInt16

        'Layer 3 used?
        If ByFlags And 4 Then .Graphic(3) = MapReader.ReadInt16

        'Layer 4 used?
        If ByFlags And 8 Then .Graphic(4) = MapReader.ReadInt16

        'Trigger used?
        If ByFlags And 16 Then .Trigger = MapReader.ReadInt16

        '.inf file
        ByFlags = InfReader.ReadInt8

        If ByFlags And 1 Then
            .TileExit.Map = InfReader.ReadInt16
            .TileExit.X = InfReader.ReadInt16
            .TileExit.Y = InfReader.ReadInt16
        End If

        If ByFlags And 2 Then
            'Get and make NPC
            .NpcIndex = InfReader.ReadInt16
            NpcIndexTest = .NpcIndex
            
            If .NpcIndex > 0 Then
                .NpcIndex = OpenNPC(.NpcIndex, True, Map)

                Npclist(.NpcIndex).Orig.Map = Map
                Npclist(.NpcIndex).Orig.X = X
                Npclist(.NpcIndex).Orig.Y = Y
                Npclist(.NpcIndex).Pos.Map = Map
                Npclist(.NpcIndex).Pos.X = X
                Npclist(.NpcIndex).Pos.Y = Y

                Call MakeNPCChar(True, 0, .NpcIndex, Map, X, Y)
            End If
        End If

        If ByFlags And 4 Then
            
            'Get and make Object
            .ObjInfo.ObjIndex = InfReader.ReadInt16
            .ObjInfo.Amount = InfReader.ReadInt16
            .ObjInfo.CurrentGrhIndex = ObjData(.ObjInfo.ObjIndex).GrhIndex
            
            If ObjData(.ObjInfo.ObjIndex).ObjType = eOBJType.otResource Then
                Call AddResourceToGroup(Map, X, Y, .ObjInfo.ObjIndex)

                'Max resource qty
                '.ObjInfo.TotalQty = ObjData(.ObjInfo.ObjIndex).MaxExtractedQuantity
                .ObjInfo.PendingQty = ObjData(.ObjInfo.ObjIndex).MaxExtractedQuantity
                .ObjInfo.Resources = ObjData(.ObjInfo.ObjIndex).Resources
            End If
            
            ' TODO: Item and Object separation
            Dim Coordinates As WorldPos

            Coordinates.Map = Map
            Coordinates.X = X
            Coordinates.Y = Y
            Call ModAreas.CreateEntity(ModAreas.Pack(Map, X, Y), ENTITY_TYPE_OBJECT, Coordinates, ObjData(.ObjInfo.ObjIndex).SizeWidth, ObjData(.ObjInfo.ObjIndex).SizeHeight)
            
        End If
        
        'Se fija si la posicion es valida para un npc de agua o tierra y la guarda por separado.
        PosType = LegalNpcSpawnPos(Map, X, Y)
        If PosType > 0 Then
            ReDim Preserve MapInfo(Map).NpcSpawnPos(PosType - 1).Pos(0 To UBound(MapInfo(Map).NpcSpawnPos(PosType - 1).Pos) + 1)
            MapInfo(Map).NpcSpawnPos(PosType - 1).Pos(UBound(MapInfo(Map).NpcSpawnPos(PosType - 1).Pos)).X = X
            MapInfo(Map).NpcSpawnPos(PosType - 1).Pos(UBound(MapInfo(Map).NpcSpawnPos(PosType - 1).Pos)).Y = Y
        End If
    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error cargando mapa: " & Map & " - Pos: " & X & "," & Y & "." & Err.Description & " - NpcIndexTest: " & NpcIndexTest)
End Sub


Public Sub GenerateEmptyResources(ByVal Map As Long)
    Dim I As Integer
    Dim J As Integer
    Dim ResourcesToRemoveQty As Integer
    Dim MinToRemove As Integer
    Dim MaxToRemove As Integer
    Dim Profession As Integer
    Dim CanExtractInMap As Boolean
    
    With MapInfo(Map).MapResources
        For I = 1 To .ResourceGroupQty
            Profession = ObjData(.ResourceGroup(I).ObjNumber).ProfessionType
            CanExtractInMap = MapInfo(Map).Pk = True Or (MapInfo(Map).Pk = False And Professions(Profession).EnabledInSafeZone)
            
            ' If the profession is disabled in safe zones and the map is safe, then we skip this
            If .ResourceGroup(I).ResourceQty > 0 And CanExtractInMap Then
                

                ' How many resources do we need to remove?
                If .ResourceGroup(I).ResourceQty > 1 Then
                    MinToRemove = Int((.ResourceGroup(I).ResourceQty / 100) * Professions(Profession).MinRemovableResourcesPercent)
                    MaxToRemove = Int((.ResourceGroup(I).ResourceQty / 100) * Professions(Profession).MaxRemovableResourcesPercent)
                    
                    MinToRemove = IIf(MinToRemove > 0, MinToRemove, 1)
                    MaxToRemove = IIf(MaxToRemove > 0, MaxToRemove, 1)
                    
                    ResourcesToRemoveQty = RandomNumber(MinToRemove, MaxToRemove)
                Else
                    ResourcesToRemoveQty = 0
                End If
                
                Dim Finished As Boolean
                Dim CantRemoved As Integer
                Dim ElementToRemove As Integer
                Dim ElementRemoved As Boolean
                Dim X As Integer
                Dim Y As Integer
                
                Do While .ResourceGroup(I).EmptyResourceQty < ResourcesToRemoveQty
                    ElementToRemove = RandomNumber(1, .ResourceGroup(I).ResourceQty - CantRemoved)
                    X = .ResourceGroup(I).ResourceList(ElementToRemove).X
                    Y = .ResourceGroup(I).ResourceList(ElementToRemove).Y
                    
                    ' If we successfully removed the element, then mark it as exhausted.
                    If AddEmptyResourceToGroup(I, Map, X, Y) Then
                        With MapData(Map, X, Y).ObjInfo
                            .PendingQty = 0
                            .CurrentGrhIndex = ObjData(.ObjIndex).DepletedGrhIndex
                        End With
                    End If
                Loop
            End If
        Next I
        
    End With
End Sub

Public Function ExhaustResourceFromGroup(ByVal ObjIndex As Integer, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte) As Boolean
    Dim I As Integer
    
    For I = 1 To MapInfo(Map).MapResources.ResourceGroupQty
        If MapInfo(Map).MapResources.ResourceGroup(I).ObjNumber = ObjIndex Then
            Call AddEmptyResourceToGroup(I, Map, X, Y)
            ExhaustResourceFromGroup = True
            Exit For
        End If
    Next I
    
End Function


Public Function AddEmptyResourceToGroup(ByVal ResourceGroupIndex As Integer, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte) As Boolean
    
On Error GoTo ErrHandler:
    Dim I As Integer
    With MapInfo(Map).MapResources.ResourceGroup(ResourceGroupIndex)
    
        ' We need to check first if the element exists in the list.
        For I = 1 To .EmptyResourceQty
            ' Element exists, so return false
            If .EmptyResourcePositions(I).Map = Map And .EmptyResourcePositions(I).X = X And .EmptyResourcePositions(I).Y = Y Then
                AddEmptyResourceToGroup = False
                Exit Function
            End If
        Next I

        ReDim Preserve .EmptyResourcePositions(1 To .EmptyResourceQty + 1)
        .EmptyResourceQty = .EmptyResourceQty + 1
        
        ' Assign values to the new element
        .EmptyResourcePositions(.EmptyResourceQty).Map = Map
        .EmptyResourcePositions(.EmptyResourceQty).X = X
        .EmptyResourcePositions(.EmptyResourceQty).Y = Y

    End With
    
    AddEmptyResourceToGroup = True
    
    Exit Function
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function AddEmptyResourceToGroup de FileIO")
    AddEmptyResourceToGroup = False
End Function

Public Sub AddResourceToGroup(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal ObjNumber As Integer)
    On Error GoTo ErrHandler:

    Dim J As Integer
    Dim FoundGroup As Integer
    FoundGroup = 0
    
    ' If there's no groups, then we add this to avoid looping through an empty array
    If MapInfo(Map).MapResources.ResourceGroupQty = 0 Then
        ReDim MapInfo(Map).MapResources.ResourceGroup(1 To 1)
        MapInfo(Map).MapResources.ResourceGroupQty = 1
        
        
        With MapInfo(Map).MapResources.ResourceGroup(1)
            .ObjNumber = ObjNumber
            ReDim .ResourceList(1 To 1)
            .ResourceQty = .ResourceQty + 1
            
            ' Assign values to the new element
            .ResourceList(.ResourceQty).Map = Map
            .ResourceList(.ResourceQty).X = X
            .ResourceList(.ResourceQty).Y = Y
        End With
        Exit Sub
    End If

    ' Loop through all the groups to see if there's one group for the given obj number
    For J = 1 To MapInfo(Map).MapResources.ResourceGroupQty
        ' Found a group based on this object?
        If MapInfo(Map).MapResources.ResourceGroup(J).ObjNumber = ObjNumber Then
            FoundGroup = J
            'Call LogError("Found group: " & FoundGroup)
            Exit For
        End If
    Next J
    
    ' If no group has been found, then we create one and then we add the resource to the group
    If FoundGroup = 0 Then
        MapInfo(Map).MapResources.ResourceGroupQty = MapInfo(Map).MapResources.ResourceGroupQty + 1
        ReDim Preserve MapInfo(Map).MapResources.ResourceGroup(1 To MapInfo(Map).MapResources.ResourceGroupQty)
        'ReDim Preserve MapInfo(Map).EmptyResourceGoups(1 To MapInfo(Map).EmptyGroupsQty)
        
        'Call LogError("Group not found. Created group : " & MapInfo(Map).EmptyGroupsQty & " - Item:" & ObjResource & ", X:" & X & ", Y: " & Y)
        
        With MapInfo(Map).MapResources.ResourceGroup(MapInfo(Map).MapResources.ResourceGroupQty)
            .ObjNumber = ObjNumber
            ReDim .ResourceList(1 To 1)
            .ResourceQty = 1
            
            ' Assign values to the new element
            .ResourceList(.ResourceQty).Map = Map
            .ResourceList(.ResourceQty).X = X
            .ResourceList(.ResourceQty).Y = Y

        End With
        Exit Sub
    Else
    ' A group has been found, so we add the empty resoure position to the group.
        With MapInfo(Map).MapResources.ResourceGroup(FoundGroup)
            'Call LogError("Assigning to found group: " & FoundGroup & " - Item:" & ObjResource & ", X:" & X & ", Y: " & Y)
            ReDim Preserve .ResourceList(1 To .ResourceQty + 1)
            .ResourceQty = .ResourceQty + 1

            ' Assign values to the new element
            .ResourceList(.ResourceQty).Map = Map
            .ResourceList(.ResourceQty).X = X
            .ResourceList(.ResourceQty).Y = Y
        End With
    End If
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GroupResource de FileIO")
End Sub


Public Function LegalNpcSpawnPos(ByVal Map As Long, ByVal X As Byte, ByVal Y As Byte) As Byte
'***************************************************
'Author: Anagrama
'Last Modification: 07/01/2017
'Revisa si la posición es válida para el spawn de un npc de tierra o agua de forma diferenciada.
'Luego devuelve si es válida, si es de tierra o agua.
'***************************************************
On Error GoTo ErrHandler
    Dim IsLegal As Boolean
    With MapData(Map, X, Y)
        IsLegal = LegalPos(Map, X, Y, False, True, True)
        IsLegal = IsLegal And (.Trigger <> eTrigger.POSINVALIDA)
        IsLegal = IsLegal And InMapBounds(Map, X, Y)
        If IsLegal = True Then
            LegalNpcSpawnPos = 1
            Exit Function
        End If
        IsLegal = LegalPos(Map, X, Y, True, False, True)
        IsLegal = IsLegal And (.Trigger <> eTrigger.POSINVALIDA)
        IsLegal = IsLegal And InMapBounds(Map, X, Y)
        If IsLegal = True Then
            LegalNpcSpawnPos = 2
            Exit Function
        End If
    End With
    
    Exit Function
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function LegalNpcSpawnPos de FileIO")
End Function

Sub LoadSini()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  
    IniPath = App.Path & "\"
    
    Dim Temporal As Long
    Dim IniManager As New clsIniManager
    
    Call IniManager.Initialize(IniPath & "Server.ini")
    
    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando info de inicio del server."
    
    BootDelBackUp = Val(IniManager.GetValue("INIT", "IniciarDesdeBackUp"))
    
    'Misc
    #If EnableSecurity Then
        Call Security.SetServerIp(IniManager.GetValue("INIT", "ServerIp"))
    #End If
    
    Puerto = Val(IniManager.GetValue("INIT", "StartPort"))
    HideMe = Val(IniManager.GetValue("INIT", "Hide"))
    AllowMultiLogins = Val(IniManager.GetValue("INIT", "AllowMultiLogins"))
    IdleLimit = Val(IniManager.GetValue("INIT", "IdleLimit"))
    'Lee la version correcta del cliente
    ULTIMAVERSION = IniManager.GetValue("INIT", "Version")
    
    PuedeCrearPersonajes = Val(IniManager.GetValue("INIT", "PuedeCrearPersonajes"))
    ServerSoloGMs = Val(IniManager.GetValue("init", "ServerSoloGMs"))
    HappyHour = Val(IniManager.GetValue("init", "HappyHour"))

    ServerConfiguration.IpTablesSecurityEnabled = GetBooleanOrDefault(IniManager.GetValue("INIT", "IpTablesSecurityEnabled"), True)
    ServerConfiguration.IpTablesSecurityLogFailedEnabled = GetBooleanOrDefault(Val(IniManager.GetValue("INIT", "IpTablesSecurityEnabled")), False)
    
    lNumHappyDays = 0
    Dim lDay As Long
    For lDay = 1 To Val(IniManager.GetValue("init", "HappyDays"))
        lNumHappyDays = lNumHappyDays + 1
        HappyHourDays(lDay) = Val(IniManager.GetValue("init", "HappyDay" & lNumHappyDays))
    Next lDay
    
    ServerConfiguration.UseExternalAccountValidation = Val(IniManager.GetValue("INIT", "UseExternalAccountValidation"))

    ArmaduraImperial1 = Val(IniManager.GetValue("INIT", "ArmaduraImperial1"))
    ArmaduraImperial2 = Val(IniManager.GetValue("INIT", "ArmaduraImperial2"))
    ArmaduraImperial3 = Val(IniManager.GetValue("INIT", "ArmaduraImperial3"))
    TunicaMagoImperial = Val(IniManager.GetValue("INIT", "TunicaMagoImperial"))
    TunicaMagoImperialEnanos = Val(IniManager.GetValue("INIT", "TunicaMagoImperialEnanos"))
    ArmaduraCaos1 = Val(IniManager.GetValue("INIT", "ArmaduraCaos1"))
    ArmaduraCaos2 = Val(IniManager.GetValue("INIT", "ArmaduraCaos2"))
    ArmaduraCaos3 = Val(IniManager.GetValue("INIT", "ArmaduraCaos3"))
    TunicaMagoCaos = Val(IniManager.GetValue("INIT", "TunicaMagoCaos"))
    TunicaMagoCaosEnanos = Val(IniManager.GetValue("INIT", "TunicaMagoCaosEnanos"))
    
    VestimentaImperialHumano = Val(IniManager.GetValue("INIT", "VestimentaImperialHumano"))
    VestimentaImperialEnano = Val(IniManager.GetValue("INIT", "VestimentaImperialEnano"))
    TunicaConspicuaHumano = Val(IniManager.GetValue("INIT", "TunicaConspicuaHumano"))
    TunicaConspicuaEnano = Val(IniManager.GetValue("INIT", "TunicaConspicuaEnano"))
    ArmaduraNobilisimaHumano = Val(IniManager.GetValue("INIT", "ArmaduraNobilisimaHumano"))
    ArmaduraNobilisimaEnano = Val(IniManager.GetValue("INIT", "ArmaduraNobilisimaEnano"))
    ArmaduraGranSacerdote = Val(IniManager.GetValue("INIT", "ArmaduraGranSacerdote"))
    
    VestimentaLegionHumano = Val(IniManager.GetValue("INIT", "VestimentaLegionHumano"))
    VestimentaLegionEnano = Val(IniManager.GetValue("INIT", "VestimentaLegionEnano"))
    TunicaLobregaHumano = Val(IniManager.GetValue("INIT", "TunicaLobregaHumano"))
    TunicaLobregaEnano = Val(IniManager.GetValue("INIT", "TunicaLobregaEnano"))
    TunicaEgregiaHumano = Val(IniManager.GetValue("INIT", "TunicaEgregiaHumano"))
    TunicaEgregiaEnano = Val(IniManager.GetValue("INIT", "TunicaEgregiaEnano"))
    SacerdoteDemoniaco = Val(IniManager.GetValue("INIT", "SacerdoteDemoniaco"))
    
    MAPA_PRETORIANO = Val(IniManager.GetValue("INIT", "MapaPretoriano"))
    
    EnTesting = Val(IniManager.GetValue("INIT", "Testing"))
      
    RECORDusuarios = Val(IniManager.GetValue("INIT", "RECORD"))
      
    'Max users
    Temporal = Val(IniManager.GetValue("INIT", "MaxUsers"))
    If MaxUsers = 0 Then
        MaxUsers = Temporal
        ReDim UserList(1 To MaxUsers) As User
    End If
    
    With ServerConfiguration
    ' %%%%%%%%%%%%%%%% REMOTE SERVERS %%%%%%%%%%%%%%%%
        .ExternalTools.StateServer.Enabled = Val(IniManager.GetValue("REMOTESERVERS", "StateServerEnabled"))
        .ExternalTools.StateServer.ListenPort = CStr(IniManager.GetValue("REMOTESERVERS", "StateServerListenPort"))
        .ExternalTools.StateServer.ExePath = CStr(IniManager.GetValue("REMOTESERVERS", "StateServerExeFullPath"))
        
        .ExternalTools.ProxyServer.Enabled = Val(IniManager.GetValue("REMOTESERVERS", "ProxyServerEnabled"))
        .ExternalTools.ProxyServer.ListenPort = CStr(IniManager.GetValue("REMOTESERVERS", "ProxyServerListenPort"))
        .ExternalTools.ProxyServer.ExePath = CStr(IniManager.GetValue("REMOTESERVERS", "ProxyServerExeFullPath"))
        
        ' %%%%%%%%%%%%%%%% SESSIONS %%%%%%%%%%%%%%%%
        .Session.Lifetime = Val(IniManager.GetValue("ACCOUNTSESSIONS", "Lifetime"))
        .Session.MaxQuantity = Val(IniManager.GetValue("ACCOUNTSESSIONS", "MaxQuantity"))
        .Session.TokenSize = Val(IniManager.GetValue("ACCOUNTSESSIONS", "TokenSize"))
        
        ' %%%%%%%%%%%%%%%% RESOURCES PATHS %%%%%%%%%%%%%%%%
        .ResourcesPaths.Dats = CStr(IniManager.GetValue("RECURSOS", "Dats"))
        .ResourcesPaths.Maps = CStr(IniManager.GetValue("RECURSOS", "Maps"))
        .ResourcesPaths.WorldBackup = CStr(IniManager.GetValue("RECURSOS", "WorldBackup"))
        
        ' %%%%%%%%%%%%%%%% LOGGING %%%%%%%%%%%%%%%%
        .LogToDebuggerWindow = CBool(Val(IniManager.GetValue("LOGGING", "LogToDebuggerWindow")))
        
        .LogsPaths.GeneralPath = CStr(IniManager.GetValue("LOGGING", "LogGeneralPath"))
        .LogsPaths.DevelopmentPath = CStr(IniManager.GetValue("LOGGING", "LogDevelopmentPath"))
        .LogsPaths.GameMastersPath = CStr(IniManager.GetValue("LOGGING", "LogGameMastersPath"))
        .LogsPaths.GuildsPath = CStr(IniManager.GetValue("LOGGING", "LogGuildsPath"))
                       
        Call CreateLogsPaths
        
    End With
    
    
    If Not AreResourcesPathsSet Then
        'cannot load next data
        Exit Sub
    End If
    
    DatPath = ServerConfiguration.ResourcesPaths.Dats
    
    Ullathorpe.Map = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Mapa")
    Ullathorpe.X = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "X")
    Ullathorpe.Y = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Y")
    
    Nix.Map = GetVar(DatPath & "Ciudades.dat", "Nix", "Mapa")
    Nix.X = GetVar(DatPath & "Ciudades.dat", "Nix", "X")
    Nix.Y = GetVar(DatPath & "Ciudades.dat", "Nix", "Y")
    
    Banderbill.Map = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Mapa")
    Banderbill.X = GetVar(DatPath & "Ciudades.dat", "Banderbill", "X")
    Banderbill.Y = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Y")
    
    Lindos.Map = GetVar(DatPath & "Ciudades.dat", "Lindos", "Mapa")
    Lindos.X = GetVar(DatPath & "Ciudades.dat", "Lindos", "X")
    Lindos.Y = GetVar(DatPath & "Ciudades.dat", "Lindos", "Y")
    
    Arghal.Map = GetVar(DatPath & "Ciudades.dat", "Arghal", "Mapa")
    Arghal.X = GetVar(DatPath & "Ciudades.dat", "Arghal", "X")
    Arghal.Y = GetVar(DatPath & "Ciudades.dat", "Arghal", "Y")
    
    Arkhein.Map = GetVar(DatPath & "Ciudades.dat", "Arkhein", "Mapa")
    Arkhein.X = GetVar(DatPath & "Ciudades.dat", "Arkhein", "X")
    Arkhein.Y = GetVar(DatPath & "Ciudades.dat", "Arkhein", "Y")
    
    Nemahuak.Map = GetVar(DatPath & "Ciudades.dat", "Nemahuak", "Mapa")
    Nemahuak.X = GetVar(DatPath & "Ciudades.dat", "Nemahuak", "X")
    Nemahuak.Y = GetVar(DatPath & "Ciudades.dat", "Nemahuak", "Y")

    
    Ciudades(eCiudad.cUllathorpe) = Ullathorpe
    Ciudades(eCiudad.cNix) = Nix
    Ciudades(eCiudad.cBanderbill) = Banderbill
    Ciudades(eCiudad.cLindos) = Lindos
    Ciudades(eCiudad.cArghal) = Arghal
    Ciudades(eCiudad.cArkhein) = Arkhein
    
    ListaCiudades(eCiudad.cUllathorpe) = "Ullathorpe"
    ListaCiudades(eCiudad.cNix) = "Nix"
    ListaCiudades(eCiudad.cBanderbill) = "Banderbill"
    ListaCiudades(eCiudad.cLindos) = "Lindos"
    ListaCiudades(eCiudad.cArghal) = "Arghal"
    ListaCiudades(eCiudad.cArkhein) = "Arkhein"
    
    IdleOff = Val(IniManager.GetValue("INIT", "IdleOff"))
    
    Call LoadStartupPositions
    
    Set IniManager = Nothing
    Call MD5sCarga
    
    Set ConsultaPopular = New ConsultasPopulares
    Call ConsultaPopular.LoadData
    
#If EnableSecurity Then
    Encriptacion.StringValidacion = Encriptacion.ArmarStringValidacion
#End If

#If SeguridadTesteo Then
    Call CargarListaPermitidos
#End If

    ' Admins
    Call loadAdministrativeUsers
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadSini de FileIO.bas")
End Sub

Sub LoadStartupPositions()
    Dim I As Integer
    Dim TmpString As String
    Dim TmpArr() As String
    
    ServerConfiguration.StartPositionsQty = Val(GetVar(DatPath & "Ciudades.dat", "STARTUP", "StartPositions"))
    
    If ServerConfiguration.StartPositionsQty <= 0 Then
        ServerConfiguration.StartPositionsQty = 0
        Erase ServerConfiguration.StartPositions
        Exit Sub
    End If

    ReDim ServerConfiguration.StartPositions(1 To ServerConfiguration.StartPositionsQty)
    For I = 1 To ServerConfiguration.StartPositionsQty
        With ServerConfiguration.StartPositions(I)
            TmpString = GetVar(DatPath & "Ciudades.dat", "STARTUP", "StartPosition" & I)
        
            TmpArr = Split(TmpString, "-")
            
            If UBound(TmpArr) <> 2 Then
                MsgBox ("Error cargando Start Positions de Ciudades.dat")
                End
            End If
            
            .Map = TmpArr(0)
            .X = TmpArr(1)
            .Y = TmpArr(2)
        
        End With
    Next I

End Sub


Sub CreateLogsPaths()

    With ServerConfiguration
        If Not .LogsPaths.GeneralPath = "" And Not FileExist(.LogsPaths.GeneralPath, vbDirectory) Then
            Call MkDir(.LogsPaths.GeneralPath)
        End If
    
        If Not .LogsPaths.DevelopmentPath = "" And Not FileExist(.LogsPaths.DevelopmentPath, vbDirectory) Then
            Call MkDir(.LogsPaths.DevelopmentPath)
        End If
        
        If Not .LogsPaths.GameMastersPath = "" And Not FileExist(.LogsPaths.GameMastersPath, vbDirectory) Then
            Call MkDir(.LogsPaths.GameMastersPath)
        End If
        
        If Not .LogsPaths.GuildsPath = "" And Not FileExist(.LogsPaths.GuildsPath, vbDirectory) Then
            Call MkDir(.LogsPaths.GuildsPath)
        End If
    End With
End Sub

Sub InitResourcesPaths()
On Error GoTo ErrHandler

  frmResources.Show
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub InitResourcesPaths de FileIO.bas")
End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'Escribe VAR en un archivo
'***************************************************
On Error GoTo ErrHandler
  

writeprivateprofilestring Main, Var, Value, File
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub WriteVar de FileIO.bas")
End Sub

Sub BackUPnPc(ByVal NpcIndex As Integer, ByVal hFile As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 10/09/2010
'10/09/2010 - Pato: Optimice el BackUp de NPCs
'13/07/2016: Anagrama - Ahora guarda el escudo, casco y arma del npc.
'***************************************************
On Error GoTo ErrHandler
  

    Dim LoopC As Integer
    
    Print #hFile, "[NPC" & Npclist(NpcIndex).Numero & "]"
    
    With Npclist(NpcIndex)
        'General
        Print #hFile, "Name=" & .Name
        Print #hFile, "Desc=" & .Desc
        Print #hFile, "Head=" & Val(.Char.head)
        Print #hFile, "Body=" & Val(.Char.body)
        Print #hFile, "Heading=" & Val(.Char.heading)
        Print #hFile, "Movement=" & Val(.Movement)
        Print #hFile, "Attackable=" & Val(.Attackable)
        Print #hFile, "Comercia=" & Val(.Comercia)
        Print #hFile, "TipoItems=" & Val(.TipoItems)
        Print #hFile, "Hostil=" & Val(.Hostile)
        Print #hFile, "GiveEXP=" & Val(.GiveEXP)
        Print #hFile, "GiveGLD=" & Val(.GiveGLD)
        Print #hFile, "InvReSpawn=" & Val(.InvReSpawn)
        Print #hFile, "NpcType=" & Val(.NPCtype)
        Print #hFile, "EquippedWeapon=" & Val(.Char.WeaponAnim)
        Print #hFile, "EquippedShield=" & Val(.Char.ShieldAnim)
        Print #hFile, "EquippedHelmet=" & Val(.Char.CascoAnim)
        
        'Stats
        Print #hFile, "Alineacion=" & Val(.Stats.Alineacion)
        Print #hFile, "DEF=" & Val(.Stats.Def)
        Print #hFile, "MaxHit=" & Val(.Stats.MaxHit)
        Print #hFile, "MaxHp=" & Val(.Stats.MaxHp)
        Print #hFile, "MinHit=" & Val(.Stats.MinHit)
        Print #hFile, "MinHp=" & Val(.Stats.MinHp)
        
        'Flags
        Print #hFile, "ReSpawn=" & Val(.flags.Respawn)
        Print #hFile, "BackUp=" & Val(.flags.BackUp)
        Print #hFile, "Domable=" & Val(.flags.Domable)
        
        'Inventario
        Print #hFile, "NroItems=" & Val(.Invent.NroItems)
        If .Invent.NroItems > 0 Then
           For LoopC = 1 To .Invent.NroItems
                Print #hFile, "Obj" & LoopC & "=" & .Invent.Object(LoopC).ObjIndex & "-" & .Invent.Object(LoopC).Amount
           Next LoopC
        End If
        
        Print #hFile, ""
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub BackUPnPc de FileIO.bas")
End Sub

Sub CargarNpcBackUp(ByVal NpcIndex As Integer, ByVal NpcNumber As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    'Status
    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup Npc"
    
    Dim npcfile As String
    
    'If NpcNumber > 499 Then
    '    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
    'Else
        npcfile = DatPath & "bkNPCs.dat"
    'End If
    
    With Npclist(NpcIndex)
    
        .Numero = NpcNumber
        .Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
        .Desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
        .Movement = Val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
        .NPCtype = Val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))
        
        .Char.body = Val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
        .Char.head = Val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
        .Char.heading = Val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))
        
        .Char.ShieldAnim = Val(GetVar(npcfile, "NPC" & NpcNumber, "EquippedShield"))
        .Char.CascoAnim = Val(GetVar(npcfile, "NPC" & NpcNumber, "EquippedHelmet"))
        .Char.WeaponAnim = Val(GetVar(npcfile, "NPC" & NpcNumber, "EquippedWeapon"))
        
        .Attackable = Val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
        .Comercia = Val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
        .Hostile = Val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
        .GiveEXP = Val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP"))
        
        
        .GiveGLD = Val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD"))
        
        .InvReSpawn = Val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))
        
        .Stats.MaxHp = Val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
        .Stats.MinHp = Val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
        .Stats.MaxHit = Val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
        .Stats.MinHit = Val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
        .Stats.Def = Val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
        .Stats.Alineacion = Val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))
        
        
        
        Dim LoopC As Integer
        Dim ln As String
        .Invent.NroItems = Val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))
        If .Invent.NroItems > 0 Then
            For LoopC = 1 To MAX_INVENTORY_SLOTS
                ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
                .Invent.Object(LoopC).ObjIndex = Val(ReadField(1, ln, 45))
                .Invent.Object(LoopC).Amount = Val(ReadField(2, ln, 45))
               
            Next LoopC
        Else
            For LoopC = 1 To MAX_INVENTORY_SLOTS
                .Invent.Object(LoopC).ObjIndex = 0
                .Invent.Object(LoopC).Amount = 0
            Next LoopC
        End If
        
        .NroDrops = Val(GetVar(npcfile, "NPC" & NpcNumber, "NRODROPS"))
        If .NroDrops > 0 Then
            ReDim .Drop(1 To .NroDrops) As tDrops
            For LoopC = 1 To .NroDrops
                ln = GetVar(npcfile, "NPC" & NpcNumber, "Drop" & LoopC)
                .Drop(LoopC).DropIndex = Val(ReadField(1, ln, 45))
                .Drop(LoopC).Probabilidad = Val(ReadField(2, ln, 45))
                .Drop(LoopC).NoExcluyente = Val(ReadField(3, ln, 45))
            Next LoopC
        End If
        
        .flags.NPCActive = True
        .flags.Respawn = Val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
        .flags.BackUp = Val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
        .flags.Domable = Val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
        .flags.RespawnOrigPos = Val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))
        .flags.DistanciaMaxima = Val(GetVar(npcfile, "NPC" & NpcNumber, "DistanciaMaxima"))
        
        'Tipo de items con los que comercia
        .TipoItems = Val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CargarNpcBackUp de FileIO.bas")
End Sub

Public Sub LogBan(ByRef BannedName As String, ByRef Baneador As String, ByRef Motivo As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  
    Dim UserID As Long
    UserID = GetUserID(BannedName)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LogBan de FileIO.bas")
End Sub

Public Sub CargaApuestas()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Apuestas.Ganancias = Val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
    Apuestas.Perdidas = Val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
    Apuestas.Jugadas = Val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CargaApuestas de FileIO.bas")
End Sub


Public Function getLimit(ByVal mapa As Integer, ByVal side As Byte) As Integer
'***************************************************
'Author: Budi
'Last Modification: 31/01/2010
'Retrieves the limit in the given side in the given map.
'TODO: This should be set in the .inf map file.
'***************************************************
On Error GoTo ErrHandler
  
Dim X As Long
Dim Y As Long

If mapa <= 0 Then Exit Function

For X = 15 To 87
    For Y = 0 To 3
        Select Case side
            Case eHeading.NORTH
                getLimit = MapData(mapa, X, 7 + Y).TileExit.Map
            Case eHeading.EAST
                getLimit = MapData(mapa, 92 - Y, X).TileExit.Map
            Case eHeading.SOUTH
                getLimit = MapData(mapa, X, 94 - Y).TileExit.Map
            Case eHeading.WEST
                getLimit = MapData(mapa, 9 + Y, X).TileExit.Map
        End Select
        If getLimit > 0 Then Exit Function
    Next Y
Next X
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function getLimit de FileIO.bas")
End Function


Public Sub LoadArmadurasFaccion()
'***************************************************
'Author: ZaMa
'Last Modification: 15/04/2010
'
'***************************************************
On Error GoTo ErrHandler
  
    Dim ClassIndex As Long
    
    Dim ArmaduraIndex As Integer
    
    
    For ClassIndex = 1 To NUMCLASES
    
        ' Defensa minima para armadas altos
        ArmaduraIndex = Val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinArmyAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
        
        ' Defensa minima para armadas bajos
        ArmaduraIndex = Val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinArmyBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
        
        ' Defensa minima para caos altos
        ArmaduraIndex = Val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinCaosAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
        
        ' Defensa minima para caos bajos
        ArmaduraIndex = Val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinCaosBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
    
    
        ' Defensa media para armadas altos
        ArmaduraIndex = Val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedArmyAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
        
        ' Defensa media para armadas bajos
        ArmaduraIndex = Val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedArmyBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
        
        ' Defensa media para caos altos
        ArmaduraIndex = Val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedCaosAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
        
        ' Defensa media para caos bajos
        ArmaduraIndex = Val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedCaosBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
    
    
        ' Defensa alta para armadas altos
        ArmaduraIndex = Val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaArmyAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
        
        ' Defensa alta para armadas bajos
        ArmaduraIndex = Val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaArmyBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
        
        ' Defensa alta para caos altos
        ArmaduraIndex = Val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaCaosAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
        
        ' Defensa alta para caos bajos
        ArmaduraIndex = Val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaCaosBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
    
        ' Defensa máxima para armadas altos
        ArmaduraIndex = Val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMaxArmyAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Armada(eTipoDefArmors.ieMax) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Armada(eTipoDefArmors.ieMax) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Armada(eTipoDefArmors.ieMax) = ArmaduraIndex
        
        ' Defensa máxima para armadas bajos
        ArmaduraIndex = Val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMaxArmyBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Armada(eTipoDefArmors.ieMax) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Armada(eTipoDefArmors.ieMax) = ArmaduraIndex
        
        ' Defensa máxima para caos altos
        ArmaduraIndex = Val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMaxCaosAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Caos(eTipoDefArmors.ieMax) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Caos(eTipoDefArmors.ieMax) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Caos(eTipoDefArmors.ieMax) = ArmaduraIndex
        
        ' Defensa máxima para caos bajos
        ArmaduraIndex = Val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMaxCaosBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Caos(eTipoDefArmors.ieMax) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Caos(eTipoDefArmors.ieMax) = ArmaduraIndex

    Next ClassIndex
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadArmadurasFaccion de FileIO.bas")
End Sub

Public Sub LoadAnimations()
'***************************************************
'Author: ZaMa
'Last Modification: 11/06/2011
'
'***************************************************
On Error GoTo ErrHandler
  
    AnimHogar(eHeading.NORTH) = 44
    AnimHogar(eHeading.EAST) = 45
    AnimHogar(eHeading.SOUTH) = 46
    AnimHogar(eHeading.WEST) = 47
    
    AnimHogarNavegando(eHeading.NORTH) = 48
    AnimHogarNavegando(eHeading.EAST) = 49
    AnimHogarNavegando(eHeading.SOUTH) = 50
    AnimHogarNavegando(eHeading.WEST) = 51
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadAnimations de FileIO.bas")
End Sub

Sub LoadClasses()
'***************************************************
On Error GoTo ErrHandler

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando clases."
    
    '*****************************************************************
    'Carga la lista de objetos
    '*****************************************************************
    Dim Index As Integer
    Dim J As Integer
    Dim X As Integer
    Dim Leer As clsIniManager
    Dim ClassesQty As Integer
    Set Leer = New clsIniManager
    
    Call Leer.Initialize(DatPath & "Classes.dat")
    
    
    ClassesQty = Val(Leer.GetValue("INIT", "Classes"))
    
    frmCargando.Cargar.Min = 0
    frmCargando.Cargar.Max = ClassesQty
    frmCargando.Cargar.Value = 0
    
    
    ReDim Preserve Classes(1 To ClassesQty) As eTypeClassConfiguration
    
    
    Dim StartingItem As String
    For Index = 1 To ClassesQty

        With Classes(Index)
            .Name = Leer.GetValue("CLASS" & Index, "Name")
            .Enabled = Val(Leer.GetValue("CLASS" & Index, "Enabled"))
            
            .StartingItemsQty = Val(Leer.GetValue("CLASS" & Index, "StartingItems"))
            .StartingSpellsQty = Val(Leer.GetValue("CLASS" & Index, "StartingSpells"))
            
            If .StartingItemsQty > 0 Then
                ReDim Preserve .StartingItems(1 To .StartingItemsQty) As eTypeClassStartingItem
                For J = 1 To .StartingItemsQty
                    StartingItem = Leer.GetValue("CLASS" & Index, "StartingItem" & J)
                
                    .StartingItems(J).ItemNumber = CInt(Val(ReadField(1, StartingItem, 45)))
                    .StartingItems(J).Quantity = CInt(Val(ReadField(2, StartingItem, 45)))
                    .StartingItems(J).Equipped = CByte(Val(ReadField(3, StartingItem, 45)))
                Next J
            End If
            
            If .StartingSpellsQty > 0 Then
                ReDim Preserve .StartingSpells(1 To .StartingSpellsQty) As Integer
                For J = 1 To .StartingSpellsQty
                    .StartingSpells(J) = CInt(Leer.GetValue("CLASS" & Index, "StartingSpell" & J))
                Next J
            End If
            
            .ClassMods.BaseDamage = Val(Leer.GetValue("CLASS" & Index, "BaseDamage"))
            
            .ClassMods.DistanceDmgReduction = Val(Leer.GetValue("CLASS" & Index, "DistanceDmgReduction"))
            .ClassMods.DistanceDamageReductionStart = CByte(Val(Leer.GetValue("CLASS" & Index, "DistanceDamageReductionStart")))
            
            .ClassMods.StabChance = Val(Leer.GetValue("CLASS" & Index, "StabChance"))
            .ClassMods.StabDamageMultiplier = CDbl(Val(Leer.GetValue("CLASS" & Index, "StabDamageMultiplier")))
            
            .ClassMods.ManaPerLevelMultiplier = CSng(Val(Leer.GetValue("CLASS" & Index, "ManaPerLevelMultiplier")))
            .ClassMods.ManaStarterMultiplier = CSng(Val(Leer.GetValue("CLASS" & Index, "ManaStarterMultiplier")))
            
            .ClassMods.HealthPerLevelMin = CByte(Val(Leer.GetValue("CLASS" & Index, "HealthPerLevelMin")))
            .ClassMods.HealthPerLevelMax = CByte(Val(Leer.GetValue("CLASS" & Index, "HealthPerLevelMax")))
            
            .ClassMods.DamageWrestlingMin = CByte(Val(Leer.GetValue("CLASS" & Index, "DamageWrestlingMin")))
            .ClassMods.DamageWrestlingMax = CByte(Val(Leer.GetValue("CLASS" & Index, "DamageWrestlingMax")))
            
            
            .ClassMods.StaminaStarter = CInt(Val(Leer.GetValue("CLASS" & Index, "StaminaStarter")))
            .ClassMods.StaminaPerLevel = CInt(Val(Leer.GetValue("CLASS" & Index, "StaminaPerLevel")))
                        
            .ClassMods.SkillsStarter = CInt(Val(Leer.GetValue("CLASS" & Index, "SkillsStarter")))
            .ClassMods.SkillsPerLevel = CInt(Val(Leer.GetValue("CLASS" & Index, "SkillsPerLevel")))
                        
            .ClassMods.MagicDamageBonus = CInt(Val(Leer.GetValue("CLASS" & Index, "MagicDamageBonus")))
            .ClassMods.MagicCastPower = CInt(Val(Leer.GetValue("CLASS" & Index, "MagicCastPower")))

            .ClassMods.MaxInvokedPets = CInt(Val(Leer.GetValue("CLASS" & Index, "MaxInvokedPets")))
            .ClassMods.MaxTammedPets = CInt(Val(Leer.GetValue("CLASS" & Index, "MaxTammedPets")))
            .ClassMods.MaxActivePets = CInt(Val(Leer.GetValue("CLASS" & Index, "MaxActivePets")))

            .ClassMods.HidingChance = CInt(Val(Leer.GetValue("CLASS" & Index, "HidingChance")))
            .ClassMods.HidingDuration = CDbl(Val(Leer.GetValue("CLASS" & Index, "HidingDuration")))

            .ClassMods.StealingChance = CInt(Val(Leer.GetValue("CLASS" & Index, "StealingChance")))
            .ClassMods.StealingAmount = CDbl(Val(Leer.GetValue("CLASS" & Index, "StealingAmount")))
            
            ' This are the same mods that were previously stored in Balance.dat
            .ClassMods.Evasion = CDbl(Val(Leer.GetValue("CLASS" & Index, "Evasion")))
            .ClassMods.AtaqueArmas = CDbl(Val(Leer.GetValue("CLASS" & Index, "WeaponAttackMod")))
            .ClassMods.AtaqueProyectiles = CDbl(Val(Leer.GetValue("CLASS" & Index, "ProjectileAttackMod")))
            .ClassMods.AtaqueWrestling = CDbl(Val(Leer.GetValue("CLASS" & Index, "WrestlingAttackMod")))
            .ClassMods.PhysicalDamage = CDbl(Val(Leer.GetValue("CLASS" & Index, "PhysicalDamageMod")))
            
            
            .ClassMods.DamageWeapons = CDbl(Val(Leer.GetValue("CLASS" & Index, "WeaponDamageMod")))
            .ClassMods.DamageProjectiles = CDbl(Val(Leer.GetValue("CLASS" & Index, "ProjectileDamageMod")))
            .ClassMods.DamageWrestling = CDbl(Val(Leer.GetValue("CLASS" & Index, "WrestlingDamageMod")))
            
            .ClassMods.Escudo = CDbl(Val(Leer.GetValue("CLASS" & Index, "ShieldMod")))
            .ClassMods.Taming = CDbl(Val(Leer.GetValue("CLASS" & Index, "TamingMod")))
            .ClassMods.Work = CDbl(Val(Leer.GetValue("CLASS" & Index, "WorkMod")))
            
            ' Get the racial mods for this class.
            ReDim .RaceMods(1 To UBound(ListaRazas))
            For J = 1 To UBound(ListaRazas)
                .RaceMods(J).StartingHealth = CInt(Val(Leer.GetValue("CLASS" & Index, "Race" & ListaRazas(J) & "_StartingHealth")))
                .RaceMods(J).HealthPerLevelMin = CInt(Val(Leer.GetValue("CLASS" & Index, "Race" & ListaRazas(J) & "_HealthPerLevelMin")))
                .RaceMods(J).HealthPerLevelMax = CInt(Val(Leer.GetValue("CLASS" & Index, "Race" & ListaRazas(J) & "_HealthPerLevelMax")))
                
                ReDim .RaceMods(J).ExtraHealthAtLevel(1 To ConstantesBalance.MaxLvl)
                
                ' Now we can configure different bonuses to the health on different levels. This way we can give the character extra health
                ' when reaching certain levels
                For X = 1 To ConstantesBalance.MaxLvl
                    .RaceMods(J).ExtraHealthAtLevel(X) = CInt(Val(Leer.GetValue("CLASS" & Index, "Race" & ListaRazas(J) & "_ExtraHealthForLevel" & X)))
                Next X
            Next J
            
            ' Masteries assigned to the class
            .MasteryGroupsQty = CInt(Val(Leer.GetValue("CLASS" & Index, "MasteryGroups")))
            
            If .MasteryGroupsQty > 0 Then
                ReDim .MasteryGroups(1 To .MasteryGroupsQty)

                Dim UnsplitMasteries As String
                Dim SplitMasteries() As String
                
                For X = 1 To .MasteryGroupsQty
                    
                    UnsplitMasteries = Leer.GetValue("CLASS" & Index, "MasteryGroup" & X)
                    
                    If UnsplitMasteries <> vbNullString Then
                        
                        SplitMasteries = Split(UnsplitMasteries, "-")
                        .MasteryGroups(X).MasteriesQty = UBound(SplitMasteries) + 1
                        ReDim .MasteryGroups(X).Masteries(1 To .MasteryGroups(X).MasteriesQty)
                        
                        For J = 1 To .MasteryGroups(X).MasteriesQty
                            .MasteryGroups(X).Masteries(J) = CInt(SplitMasteries(J - 1))
                        Next J
                    
                    Else
                        .MasteryGroups(X).MasteriesQty = 0
                    End If
                
                Next X
            
            End If
            
        End With
        
        frmCargando.Cargar.Value = frmCargando.Cargar.Value + 1
    Next Index
    
    
    Set Leer = Nothing

    Exit Sub

ErrHandler:
    MsgBox "error cargando clase " & Index & ": " & Err.Number & ": " & Err.Description

End Sub

Public Function AreResourcesPathsSet() As Boolean
    With ServerConfiguration.ResourcesPaths
        If Not FileExist(.Dats, vbDirectory) Or Trim(.Dats) = "" Then
            AreResourcesPathsSet = False
            Exit Function
        End If
         If Not FileExist(.Maps, vbDirectory) Or Trim(.Maps) = "" Then
            AreResourcesPathsSet = False
            Exit Function
        End If
        If Not FileExist(.WorldBackup, vbDirectory) Or Trim(.WorldBackup) = "" Then
            AreResourcesPathsSet = False
            Exit Function
        End If
    End With
        
    AreResourcesPathsSet = True
    Exit Function
End Function

Public Sub LoadMasteries()
    On Error GoTo ErrHandler

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando maestrías."
    
    Dim Index As Integer
    Dim J As Integer
    Dim Leer As clsIniManager
    Set Leer = New clsIniManager
    
    Call Leer.Initialize(ServerConfiguration.ResourcesPaths.Dats & "Masteries.dat")
    
    
    MasteriesQty = Val(Leer.GetValue("INIT", "Masteries"))
    
    If MasteriesQty = 0 Then
        Exit Sub
    End If
    
    frmCargando.Cargar.Min = 0
    frmCargando.Cargar.Max = MasteriesQty
    frmCargando.Cargar.Value = 0
        
    ReDim Preserve Masteries(1 To MasteriesQty)
        
    Dim StartingItem As String
    For Index = 1 To MasteriesQty
        
        With Masteries(Index)
            .Id = Index
            .Name = Leer.GetValue("MASTERY" & Index, "Name")
            .Description = Leer.GetValue("MASTERY" & Index, "Description")
            .Enabled = Val(Leer.GetValue("MASTERY" & Index, "Enabled"))
                        
            .GoldRequired = Val(Leer.GetValue("MASTERY" & Index, "GoldRequired"))
            .PointsRequired = Val(Leer.GetValue("MASTERY" & Index, "PointsRequired"))
            .MasteryRequired = Val(Leer.GetValue("MASTERY" & Index, "MasteryRequired"))
            
            ' Additive masteries.
            .MagicSpellDamageLeechPerc = Val(Leer.GetValue("MASTERY" & Index, "MagicSpellDamageLeechPerc"))
            .MagicSpellManaConversionPerc = Val(Leer.GetValue("MASTERY" & Index, "MagicSpellManaConversionPerc"))

            .AddTamingPoints = Val(Leer.GetValue("MASTERY" & Index, "AddTamingPoints"))
                                    
            ' Masteries that impacts based on a list of criterias
            Dim TempValue As Integer
            Dim X As Integer
                                                    
            ' CanDissarmWithItem list calculation
            .CanDissarmWithItemQty = Val(Leer.GetValue("MASTERY" & Index, "CanDissarmWithItemQty"))
            If .CanDissarmWithItemQty > 0 Then
                ReDim Preserve .CanDissarmWithItem(1 To .CanDissarmWithItemQty)
                For X = 1 To .CanDissarmWithItemQty
                    .CanDissarmWithItem(X) = Val(Leer.GetValue("MASTERY" & Index, "CanDissarmWithItem" & X))
                Next X
            End If
            
            ' -==== Finished masteries ====-
            .EnableBerserkWhileSailing = CBool(Val(Leer.GetValue("MASTERY" & Index, "EnableBerserkWhileSailing")))
            .AddMagicLifeLeechPerc = Val(Leer.GetValue("MASTERY" & Index, "AddMagicLifeLeechPerc"))
            .AddMagicCastPower = Val(Leer.GetValue("MASTERY" & Index, "AddMagicCastPower"))
            .AddInviMinDuration = Val(Leer.GetValue("MASTERY" & Index, "AddInviMinDuration"))
            .AddInviMaxDuration = Val(Leer.GetValue("MASTERY" & Index, "AddInviMaxDuration"))
            .AddStabChanceWhenInviPerc = Val(Leer.GetValue("MASTERY" & Index, "AddStabChanceWhenInviPerc"))
            .AddBackstabDamageBonusPerc = Val(Leer.GetValue("MASTERY" & Index, "AddBackstabDamageBonusPerc"))
            
            .AddMaxHealth = Val(Leer.GetValue("MASTERY" & Index, "AddMaxHealth"))
            .AddMaxMana = Val(Leer.GetValue("MASTERY" & Index, "AddMaxMana"))
            .AddMaxManaPerc = Val(Leer.GetValue("MASTERY" & Index, "AddMaxManaPerc"))
            
            .AddBaseWeaponDamagePercent = Val(Leer.GetValue("MASTERY" & Index, "AddBaseWeaponDamagePercent"))
            .AddBaseWrestlingDamagePercent = Val(Leer.GetValue("MASTERY" & Index, "AddBaseWrestlingDamagePercent"))
            .AddBaseRangedDamagePercent = Val(Leer.GetValue("MASTERY" & Index, "AddBaseRangedDamagePercent"))
            .AddBaseMagicDamagePercent = Val(Leer.GetValue("MASTERY" & Index, "AddBaseMagicDamagePercent"))
            .AddEnergyRegeneration = Val(Leer.GetValue("MASTERY" & Index, "AddEnergyRegeneration"))
            .AddExtraHitChance = Val(Leer.GetValue("MASTERY" & Index, "AddExtraHitChance"))
            
            
            Dim TmpValues() As String
            
            ' BypassProhibitedClasses list calculation
            .BypassProhibitedClassesQty = Val(Leer.GetValue("MASTERY" & Index, "BypassProhibitedClassesQty"))
            If .BypassProhibitedClassesQty > 0 Then
                ReDim Preserve .BypassProhibitedClassesObjs(1 To .BypassProhibitedClassesQty)
                For X = 1 To .BypassProhibitedClassesQty
                    .BypassProhibitedClassesObjs(X) = Val(Leer.GetValue("MASTERY" & Index, "BypassProhibitedClassesObj" & X))
                Next X
            End If
            
            ' Mana Cost Reduction for spells
            .SpellManaCostReductionQty = Val(Leer.GetValue("MASTERY" & Index, "SpellManaCostReductionQty"))
            If .SpellManaCostReductionQty > 0 Then
                ReDim Preserve .SpellManaCostReduction(1 To .SpellManaCostReductionQty)
                For X = 1 To .SpellManaCostReductionQty
                    ' Format of this property value is SpellNumber-PercentOfManaReduction
                    TmpValues = Split(Leer.GetValue("MASTERY" & Index, "SpellManaCostReduction" & X), "-")
                    
                    .SpellManaCostReduction(X).Spell = CInt(TmpValues(0))
                    .SpellManaCostReduction(X).ValuePercent = CInt(TmpValues(1))
                Next X
            End If
            
            ' MagicBonusForSpell list calculation
            .MagicBonusForSpellQty = Val(Leer.GetValue("MASTERY" & Index, "MagicBonusForSpellQty"))
            If .MagicBonusForSpellQty > 0 Then
                ReDim Preserve .MagicBonusForSpell(1 To .MagicBonusForSpellQty)
                For X = 1 To .MagicBonusForSpellQty
                    ' Format of this property value is SpellNumber-PercentOfManaReduction
                    TmpValues = Split(Leer.GetValue("MASTERY" & Index, "MagicBonusForSpell" & X), "-")
                    
                    .MagicBonusForSpell(X).Spell = CInt(TmpValues(0))
                    .MagicBonusForSpell(X).ValuePercent = CInt(TmpValues(1))
                Next X
            End If
            
            ' ImmunityToSpell list calculation
            .ImmunityToSpellQty = Val(Leer.GetValue("MASTERY" & Index, "ImmunityToSpellQty"))
            If .ImmunityToSpellQty > 0 Then
                ReDim Preserve .ImmunityToSpell(1 To .ImmunityToSpellQty)
                For X = 1 To .ImmunityToSpellQty
                    .ImmunityToSpell(X) = CInt(Leer.GetValue("MASTERY" & Index, "ImmunityToSpell" & X))
                Next X
            End If
             
        
             ' -==== / Finished masteries ====-
        End With

        frmCargando.Cargar.Value = frmCargando.Cargar.Value + 1
    Next Index
    
    
    Set Leer = Nothing

    Exit Sub

ErrHandler:
    MsgBox "Error cargando maestria " & Index & ": " & Err.Number & ": " & Err.Description
End Sub


Sub LoadPassiveSkillsConfig()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim I As Integer
    Dim J As Integer
    Dim Leer As clsIniManager
    Set Leer = New clsIniManager
    
    Call Leer.Initialize(ServerConfiguration.ResourcesPaths.Dats & "PassiveSkills.dat")
    
    ServerConfiguration.PassiveSkillsQty = Val(Leer.GetValue("INIT", "PassiveQty"))
    
    If ServerConfiguration.PassiveSkillsQty <= 0 Then
        Exit Sub
    End If
    
    ReDim ServerConfiguration.PassiveSkills(1 To ServerConfiguration.PassiveSkillsQty)
    
    For I = 1 To ServerConfiguration.PassiveSkillsQty
    
        With ServerConfiguration.PassiveSkills(I)
            .Name = Leer.GetValue("PASSIVE" & I, "Name")
            .Enabled = Val(Leer.GetValue("PASSIVE" & I, "Enabled"))
            .UnlockLevel = Val(Leer.GetValue("PASSIVE" & I, "UnlockLevel"))
            
            .AllowedClassesQty = CInt(Val(Leer.GetValue("PASSIVE" & I, "AllowedClassesQty")))
            If .AllowedClassesQty > 0 Then
                ReDim .AllowedClasses(1 To .AllowedClassesQty)
                
                For J = 1 To .AllowedClassesQty
                    .AllowedClasses(J) = GetClassTypeFromName(Leer.GetValue("PASSIVE" & I, "AllowedClasses" & J))
                Next J
                
            Else
                Erase .AllowedClasses
            End If
            
        End With
        
    Next I
    
    
    Set Leer = Nothing
    Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadPassiveSkillsConfig de FileIO.bas")
End Sub

Public Function GetBooleanOrDefault(ByRef Value As String, ByVal DefaultValue As Boolean)
    If Value = vbNullString Or Not IsNumeric(Value) Then
        GetBooleanOrDefault = DefaultValue
    Else
        GetBooleanOrDefault = CBool(Value)
    End If
End Function
