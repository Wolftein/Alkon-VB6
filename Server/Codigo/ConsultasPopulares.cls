VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConsultasPopulares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' ConsultasPopulares.cls
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
'Modulo de consultas popluares
'En realidad el modulo inicial guardaba los datos de las votaciones
'en memoria, pero no le vi el punto, las votaciones son de vez en cuando nomas
'y el query a un .ini que registra todos los mails q ya votaron, es soportable
'si cuando uno vota y averiguo que el mail ya voto pero el PJ no, entonces seteo
'el flag de yavoto en el charfile ese tambien,
'por lo que la busqueda secuencial en el .dat que tiene todos los mails q ya votaron
'se hara .. 1 vez por PJ nomas.
'
'Hecha por el oso

Private Const ARCHIVOMAILS = "\logs\votaron.dat"
Private Const ARCHIVOCONFIG = "consultas.dat"

Private pEncuestaActualNum As Integer
Private pEncuestaActualTex As String
Private pNivelRequerido As Integer
Private pOpciones() As Integer


Public Property Let Numero(ByVal NumEncuesta As Integer)
    pEncuestaActualNum = NumEncuesta
End Property

Public Property Get Numero() As Integer
    Numero = pEncuestaActualNum
End Property

Public Property Let texto(ByVal Descripcion As String)
    pEncuestaActualTex = Descripcion
End Property

Public Property Get texto() As String
    texto = pEncuestaActualTex
End Property


Public Sub LoadData()
On Error GoTo ErrHandler
  
Dim CantOpciones As Integer
Dim I As Integer

    pEncuestaActualNum = Val(GetVar(ServerConfiguration.ResourcesPaths.Dats & ARCHIVOCONFIG, "INIT", "ConsultaActual"))
    pEncuestaActualTex = GetVar(ServerConfiguration.ResourcesPaths.Dats & ARCHIVOCONFIG, "INIT", "ConsultaActualTexto")
    pNivelRequerido = GetVar(ServerConfiguration.ResourcesPaths.Dats & ARCHIVOCONFIG, "INIT", "NivelRequerido")

    If pEncuestaActualNum > 0 Then
       'cargo todas las opciones
        CantOpciones = Val(GetVar(ServerConfiguration.ResourcesPaths.Dats & ARCHIVOCONFIG, "ENCUESTA" & pEncuestaActualNum, "CANTOPCIONES"))
        ReDim pOpciones(1 To CantOpciones) As Integer
        For I = 1 To CantOpciones
            pOpciones(I) = Val(GetVar(ServerConfiguration.ResourcesPaths.Dats & ARCHIVOCONFIG, "ENCUESTA" & pEncuestaActualNum, "OPCION" & I))
        Next I
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadData de ConsultasPopulares.cls")
End Sub

Public Function doVotar(ByVal UserIndex As Integer, ByVal opcion As Integer) As String
On Error GoTo errorh

    'revisar q no haya votado
    'grabar en el charfile el numero de encuesta
    'actualizar resultados encuesta
    If pEncuestaActualNum = 0 Then
        doVotar = "No hay consultas populares abiertas"
        Exit Function
    End If

    Dim CharFile As String
    CharFile = CharPath & UserList(UserIndex).Name & ".chr"
    
    ' Nivel valido?
    If (UserList(UserIndex).Stats.ELV < pNivelRequerido) Then
        doVotar = "Para votar en esta consulta debes ser nivel " & pNivelRequerido & " o superior"
        Exit Function
    End If

    ' Opcion valida?
    If Not (OpcionValida(opcion)) Then
        doVotar = "Esa no es una opción válida para votar"
        Exit Function
    End If
    
    
    Dim YaVoto As Boolean
    Dim YaVotoMail As Boolean
    
    ' Ya voto?
    ' Get info from DB, returned byref
    'Call GetUserPollInfo(UserList(UserIndex).ID, pEncuestaActualNum, UserList(UserIndex).Email, YaVoto, YaVotoMail)

    If YaVoto Then
        doVotar = "Este personaje ya voto en la consulta: " & pEncuestaActualTex
        Exit Function
    End If
    
    ' Mail ya voto?

    If YaVotoMail Then
        doVotar = "Este email ya voto en la consulta: " & pEncuestaActualTex
        Exit Function
    End If
    
    'pj apto para votar
    Dim sufragio As Long
    sufragio = CLng(Val(GetVar(App.Path & ARCHIVOCONFIG, "RESULTADOS" & pEncuestaActualNum, "V" & opcion)))
    sufragio = sufragio + 1
    Call WriteVar(App.Path & ARCHIVOCONFIG, "RESULTADOS" & pEncuestaActualNum, "V" & opcion, str(sufragio))

    'Call SaveUserVoteDB(UserList(UserIndex).ID, pEncuestaActualNum, opcion, UserList(UserIndex).Email)

    doVotar = "Tu voto ha sido computado. Opcion: " & opcion

    Exit Function
errorh:
    Call LogError("Error en ConsultasPopularse.doVotar: " & Err.Description)
End Function


Public Function SendInfoEncuesta(ByVal UserIndex As Integer) As String
On Error GoTo ErrHandler
  
Dim I As Integer
    Call WriteConsoleMsg(UserIndex, "CONSULTA POPULAR NUMERO " & pEncuestaActualNum, FontTypeNames.FONTTYPE_GUILD)
    Call WriteConsoleMsg(UserIndex, pEncuestaActualTex, FontTypeNames.FONTTYPE_GUILD)
    Call WriteConsoleMsg(UserIndex, " Opciones de voto: ", FontTypeNames.FONTTYPE_GUILDMSG)
    For I = 1 To UBound(pOpciones)
        Call WriteConsoleMsg(UserIndex, "(Opcion " & I & "): " & GetVar(App.Path & ARCHIVOCONFIG, "ENCUESTA" & pEncuestaActualNum, "OPCION" & I), FontTypeNames.FONTTYPE_GUILDMSG)
    Next I
    Call WriteConsoleMsg(UserIndex, " Para votar una opcion, escribe /encuesta NUMERODEOPCION, por ejemplo para votar la opcion 1, escribe /encuesta 1. Tu voto no podra ser cambiado.", FontTypeNames.FONTTYPE_VENENO)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function SendInfoEncuesta de ConsultasPopulares.cls")
End Function

Private Sub MarcarPjComoQueYaVoto(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
  
    Call WriteVar(CharPath & UserList(UserIndex).Name & ".chr", "CONSULTAS", "Voto", str(pEncuestaActualNum))
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub MarcarPjComoQueYaVoto de ConsultasPopulares.cls")
End Sub


Private Function MailYaVoto(ByVal Email As String) As Boolean
'abro el archivo, while not eof levnato 1 linea y comparo. Si da true, cierro
On Error GoTo ErrHandler
  
Dim ArchN As Integer
Dim Tmp As String

MailYaVoto = False

ArchN = FreeFile()

Open App.Path & ARCHIVOMAILS For Input As #ArchN

Do While Not EOF(ArchN)
    Line Input #ArchN, Tmp
    If Email = Tmp Then
        MailYaVoto = True
        Close #ArchN
        Exit Function
    End If
Loop

Close #ArchN
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function MailYaVoto de ConsultasPopulares.cls")
End Function


Private Sub MarcarMailComoQueYaVoto(ByVal Email As String)
On Error GoTo ErrHandler
  
Dim ArchN As Integer

ArchN = FreeFile()

Open App.Path & ARCHIVOMAILS For Append As #ArchN
Print #ArchN, Email

Close #ArchN

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub MarcarMailComoQueYaVoto de ConsultasPopulares.cls")
End Sub

Private Function OpcionValida(ByVal opcion As Integer) As Boolean
On Error GoTo ErrHandler
  
    OpcionValida = opcion > 0 And opcion <= UBound(pOpciones)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function OpcionValida de ConsultasPopulares.cls")
End Function
