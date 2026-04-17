Attribute VB_Name = "modPrivateMessages"
Option Explicit

Public Sub AgregarMensaje(ByVal UserIndex As Integer, ByRef Autor As String, ByRef Mensaje As String)
'***************************************************
'Author: Amraphen
'Last Modification: 04/08/2011
'Agrega un nuevo mensaje privado a un usuario online.
'***************************************************
On Error GoTo ErrHandler
  
Dim LoopC As Long

    With UserList(UserIndex)
        If .UltimoMensaje < Constantes.MaxPrivateMessages Then
            .UltimoMensaje = .UltimoMensaje + 1
        Else
            For LoopC = 1 To Constantes.MaxPrivateMessages - 1
                .Mensajes(LoopC) = .Mensajes(LoopC + 1)
            Next
        End If
        
        With .Mensajes(.UltimoMensaje)
            .Contenido = UCase$(Autor) & ": " & Mensaje & " (" & Now & ")"
            .Nuevo = True
        End With
        
        Call WriteConsoleMsg(UserIndex, "¡Has recibido un mensaje privado de un Game Master!", FontTypeNames.FONTTYPE_GM)
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AgregarMensaje de modPrivateMessages.bas")
End Sub

Public Sub AgregarMensajeOFF(ByRef Destinatario As String, ByRef Autor As String, ByRef Mensaje As String)
'***************************************************
'Author: Amraphen
'Last Modification: 04/08/2011
'Agrega un nuevo mensaje privado a un usuario offline.
'***************************************************
On Error GoTo ErrHandler
  
Dim UltimoMensaje As Byte
Dim CharFile As String
Dim Contenido As String
Dim LoopC As Long

    CharFile = CharPath & Destinatario & ".chr"
    UltimoMensaje = CByte(GetVar(CharFile, "MENSAJES", "UltimoMensaje"))
    Contenido = UCase$(Autor) & ": " & Mensaje & " (" & Now & ")"

    If UltimoMensaje < Constantes.MaxPrivateMessages Then
        UltimoMensaje = UltimoMensaje + 1
    Else
        For LoopC = 1 To Constantes.MaxPrivateMessages - 1
            Call WriteVar(CharFile, "MENSAJES", "MSJ" & LoopC, GetVar(CharFile, "MENSAJES", "MSJ" & LoopC + 1))
            Call WriteVar(CharFile, "MENSAJES", "MSJ" & LoopC & "_NUEVO", GetVar(CharFile, "MENSAJES", "MSJ" & LoopC + 1 & "_NUEVO"))
        Next LoopC
    End If
        
    Call WriteVar(CharFile, "MENSAJES", "MSJ" & UltimoMensaje, Contenido)
    Call WriteVar(CharFile, "MENSAJES", "MSJ" & UltimoMensaje & "_NUEVO", 1)
    
    Call WriteVar(CharFile, "MENSAJES", "UltimoMensaje", UltimoMensaje)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AgregarMensajeOFF de modPrivateMessages.bas")
End Sub

Public Function TieneMensajesNuevos(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Amraphen
'Last Modification: 04/08/2011
'Determina si el usuario tiene mensajes nuevos.
'***************************************************
On Error GoTo ErrHandler
  
Dim LoopC As Long

    For LoopC = 1 To Constantes.MaxPrivateMessages
        If UserList(UserIndex).Mensajes(LoopC).Nuevo Then
            TieneMensajesNuevos = True
            Exit Function
        End If
    Next LoopC
    
    TieneMensajesNuevos = False
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function TieneMensajesNuevos de modPrivateMessages.bas")
End Function

Public Sub GuardarMensajes(ByVal UserIndex As Integer, ByRef Manager As clsIniManager)
'***************************************************
'Author: Amraphen
'Last Modification: 04/08/2011
'Guarda los mensajes del usuario.
'***************************************************
On Error GoTo ErrHandler
  
Dim LoopC As Long
    
    With UserList(UserIndex)
        Call Manager.ChangeValue("MENSAJES", "UltimoMensaje", CStr(.UltimoMensaje))
        
        For LoopC = 1 To Constantes.MaxPrivateMessages
            Call Manager.ChangeValue("MENSAJES", "MSJ" & LoopC, .Mensajes(LoopC).Contenido)
            If .Mensajes(LoopC).Nuevo Then
                Call Manager.ChangeValue("MENSAJES", "MSJ" & LoopC & "_NUEVO", 1)
            Else
                Call Manager.ChangeValue("MENSAJES", "MSJ" & LoopC & "_NUEVO", 0)
            End If
        Next LoopC
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub GuardarMensajes de modPrivateMessages.bas")
End Sub

Public Sub CargarMensajes(ByVal UserIndex As Integer, ByRef Manager As clsIniManager)
'***************************************************
'Author: Amraphen
'Last Modification: 04/08/2011
'Carga los mensajes del usuario.
'***************************************************
On Error GoTo ErrHandler
  
Dim LoopC As Long

    With UserList(UserIndex)
        .UltimoMensaje = Val(Manager.GetValue("MENSAJES", "UltimoMensaje"))
        
        For LoopC = 1 To Constantes.MaxPrivateMessages
            With .Mensajes(LoopC)
                .Nuevo = Val(Manager.GetValue("MENSAJES", "MSJ" & LoopC & "_NUEVO"))
                .Contenido = CStr(Manager.GetValue("MENSAJES", "MSJ" & LoopC))
            End With
        Next LoopC
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CargarMensajes de modPrivateMessages.bas")
End Sub

Private Sub LimpiarMensajeSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Amraphen
'Last Modification: 04/08/2011
'Limpia el un mensaje de un usuario online.
'***************************************************
On Error GoTo ErrHandler
  
    With UserList(UserIndex).Mensajes(Slot)
        .Contenido = vbNullString
        .Nuevo = False
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LimpiarMensajeSlot de modPrivateMessages.bas")
End Sub

Public Sub LimpiarMensajes(ByVal UserIndex As Integer)
'***************************************************
'Author: Amraphen
'Last Modification: 04/08/2011
'Limpia los mensajes del slot.
'***************************************************
On Error GoTo ErrHandler
  
Dim LoopC As Long

    With UserList(UserIndex)
        .UltimoMensaje = 0
        
        For LoopC = 1 To Constantes.MaxPrivateMessages
            Call LimpiarMensajeSlot(UserIndex, LoopC)
        Next LoopC
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LimpiarMensajes de modPrivateMessages.bas")
End Sub

Public Sub BorrarMensaje(ByVal UserIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Amraphen
'Last Modification: 04/08/2011
'Borra un mensaje de un usuario.
'***************************************************
On Error GoTo ErrHandler
  
Dim LoopC As Long

    With UserList(UserIndex)
        If Slot > .UltimoMensaje Or Slot < 1 Then Exit Sub

        If Slot = .UltimoMensaje Then
            Call LimpiarMensajeSlot(UserIndex, Slot)
        Else
            For LoopC = Slot To Constantes.MaxPrivateMessages - 1
                .Mensajes(LoopC) = .Mensajes(LoopC + 1)
            Next LoopC
            Call LimpiarMensajeSlot(UserIndex, .UltimoMensaje)
        End If
        
        .UltimoMensaje = .UltimoMensaje - 1
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub BorrarMensaje de modPrivateMessages.bas")
End Sub

Public Sub BorrarMensajeOFF(ByVal UserName As String, ByVal Slot As Byte)
'***************************************************
'Author: Amraphen
'Last Modification: 04/08/2011
'Borra un mensaje de un usuario.
'***************************************************
On Error GoTo ErrHandler
  
Dim CharFile As String
Dim UltimoMensaje As Byte
Dim LoopC As Long

    CharFile = CharPath & UserName & ".chr"
    
    UltimoMensaje = GetVar(CharFile, "MENSAJES", "UltimoMensaje")
    
    If Slot > UltimoMensaje Or Slot < 1 Then Exit Sub
    
    If Slot = UltimoMensaje Then
        Call WriteVar(CharFile, "MENSAJES", "MSJ" & Slot, vbNullString)
        Call WriteVar(CharFile, "MENSAJES", "MSJ" & Slot & "_Nuevo", vbNullString)
    Else
        For LoopC = Slot To UltimoMensaje - 1
            Call WriteVar(CharFile, "MENSAJES", "MSJ" & LoopC, GetVar(CharFile, "MENSAJES", "MSJ" & LoopC + 1))
            Call WriteVar(CharFile, "MENSAJES", "MSJ" & LoopC & "_NUEVO", GetVar(CharFile, "MENSAJES", "MSJ" & LoopC + 1 & "_NUEVO"))
        Next LoopC
        Call WriteVar(CharFile, "MENSAJES", "MSJ" & UltimoMensaje, vbNullString)
        Call WriteVar(CharFile, "MENSAJES", "MSJ" & UltimoMensaje & "_Nuevo", vbNullString)
    End If
    
    Call WriteVar(CharFile, "MENSAJES", "UltimoMensaje", UltimoMensaje - 1)
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub BorrarMensajeOFF de modPrivateMessages.bas")
End Sub

Public Sub LimpiarMensajesOFF(ByVal UserName As String)
'***************************************************
'Author: Amraphen
'Last Modification: 18/08/2011
'Borra los mensajes de un usuario offline.
'***************************************************
On Error GoTo ErrHandler
  
Dim CharFile As String
Dim UltimoMensaje As Byte
Dim LoopC As Long

    CharFile = CharPath & UserName & ".chr"
    
    UltimoMensaje = GetVar(CharFile, "MENSAJES", "UltimoMensaje")
    
    If UltimoMensaje > 0 Then
        For LoopC = 1 To UltimoMensaje
            Call WriteVar(CharFile, "MENSAJES", "MSJ" & LoopC, vbNullString)
            Call WriteVar(CharFile, "MENSAJES", "MSJ" & LoopC & "_NUEVO", vbNullString)
        Next LoopC
        
        Call WriteVar(CharFile, "MENSAJES", "UltimoMensaje", 0)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LimpiarMensajesOFF de modPrivateMessages.bas")
End Sub

Public Sub EnviarMensaje(ByVal UserIndex As Integer, ByVal MpIndex As Byte, ByRef Mensaje As String, _
    ByVal bNuevo As Boolean)
'***************************************************
'Author: Amraphen
'Last Modification: 18/08/2011
'Envía mensaje por consola.
'***************************************************
On Error GoTo ErrHandler
  
    If LenB(Mensaje) Then
        If bNuevo Then
            Call WriteConsoleMsg(UserIndex, "MENSAJE " & MpIndex & "> (!) " & Mensaje, FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "MENSAJE " & MpIndex & "> " & Mensaje, FontTypeNames.FONTTYPE_INFO)
        End If
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EnviarMensaje de modPrivateMessages.bas")
End Sub
