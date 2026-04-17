Attribute VB_Name = "modUserRecords"
'Argentum Online 0.13.0
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

Public Sub LoadRecords()
'**************************************************************
'Author: Amraphen
'Last Modify Date: 29/11/2010
'Carga los seguimientos de usuarios.
'**************************************************************
On Error GoTo ErrHandler
  
Dim Reader As clsIniManager
Dim TmpStr As String
Dim I As Long
Dim J As Long

    Set Reader = New clsIniManager
    
    If Not FileExist(DatPath & "RECORDS.DAT") Then
        Call CreateRecordsFile
    End If
    
    Call Reader.Initialize(DatPath & "RECORDS.DAT")

    NumRecords = Reader.GetValue("INIT", "NumRecords")
    If NumRecords Then ReDim Records(1 To NumRecords)
    
    For I = 1 To NumRecords
        With Records(I)
            .Usuario = Reader.GetValue("RECORD" & I, "Usuario")
            .Creador = Reader.GetValue("RECORD" & I, "Creador")
            .Fecha = Reader.GetValue("RECORD" & I, "Fecha")
            .Motivo = Reader.GetValue("RECORD" & I, "Motivo")

            .NumObs = Val(Reader.GetValue("RECORD" & I, "NumObs"))
            If .NumObs Then ReDim .Obs(1 To .NumObs)
            
            For J = 1 To .NumObs
                TmpStr = Reader.GetValue("RECORD" & I, "Obs" & J)
                
                .Obs(J).Creador = ReadField(1, TmpStr, 45)
                .Obs(J).Fecha = ReadField(2, TmpStr, 45)
                .Obs(J).Detalles = ReadField(3, TmpStr, 45)
            Next J
        End With
    Next I
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadRecords de modUserRecords.bas")
End Sub

Public Sub SaveRecords()
'**************************************************************
'Author: Amraphen
'Last Modify Date: 29/11/2010
'Guarda los seguimientos de usuarios.
'**************************************************************
On Error GoTo ErrHandler
  
Dim Writer As clsIniManager
Dim TmpStr As String
Dim I As Long
Dim J As Long

    Set Writer = New clsIniManager

    Call Writer.ChangeValue("INIT", "NumRecords", NumRecords)
    
    For I = 1 To NumRecords
        With Records(I)
            Call Writer.ChangeValue("RECORD" & I, "Usuario", .Usuario)
            Call Writer.ChangeValue("RECORD" & I, "Creador", .Creador)
            Call Writer.ChangeValue("RECORD" & I, "Fecha", .Fecha)
            Call Writer.ChangeValue("RECORD" & I, "Motivo", .Motivo)
            
            Call Writer.ChangeValue("RECORD" & I, "NumObs", .NumObs)
            
            For J = 1 To .NumObs
                TmpStr = .Obs(J).Creador & "-" & .Obs(J).Fecha & "-" & .Obs(J).Detalles
                Call Writer.ChangeValue("RECORD" & I, "Obs" & J, TmpStr)
            Next J
        End With
    Next I
    
    Call Writer.DumpFile(DatPath & "RECORDS.DAT")
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SaveRecords de modUserRecords.bas")
End Sub

Public Sub AddRecord(ByVal UserIndex As Integer, ByVal Nickname As String, ByVal Reason As String)
'**************************************************************
'Author: Amraphen
'Last Modify Date: 29/11/2010
'Agrega un seguimiento.
'**************************************************************
On Error GoTo ErrHandler
  
    NumRecords = NumRecords + 1
    ReDim Preserve Records(1 To NumRecords)
    
    With Records(NumRecords)
        .Usuario = UCase$(Nickname)
        .Fecha = Format(Now, "DD/MM/YYYY hh:mm:ss")
        .Creador = UCase$(UserList(UserIndex).Name)
        .Motivo = Reason
        .NumObs = 0
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AddRecord de modUserRecords.bas")
End Sub

Public Sub AddObs(ByVal UserIndex As Integer, ByVal RecordIndex As Integer, ByVal Obs As String)
'**************************************************************
'Author: Amraphen
'Last Modify Date: 29/11/2010
'Agrega una observación.
'**************************************************************
On Error GoTo ErrHandler
  
    With Records(RecordIndex)
        .NumObs = .NumObs + 1
        ReDim Preserve .Obs(1 To .NumObs)
        
        .Obs(.NumObs).Creador = UCase$(UserList(UserIndex).Name)
        .Obs(.NumObs).Fecha = Now
        .Obs(.NumObs).Detalles = Obs
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AddObs de modUserRecords.bas")
End Sub

Public Sub RemoveRecord(ByVal RecordIndex As Integer)
'**************************************************************
'Author: Amraphen
'Last Modify Date: 29/11/2010
'Elimina un seguimiento.
'**************************************************************
On Error GoTo ErrHandler
  
Dim I As Long
    
    If RecordIndex = NumRecords Then
        NumRecords = NumRecords - 1
        If NumRecords > 0 Then
            ReDim Preserve Records(1 To NumRecords)
        End If
    Else
        NumRecords = NumRecords - 1
        For I = RecordIndex To NumRecords
            Records(I) = Records(I + 1)
        Next I

        ReDim Preserve Records(1 To NumRecords)
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RemoveRecord de modUserRecords.bas")
End Sub

Public Sub CreateRecordsFile()
'**************************************************************
'Author: Amraphen
'Last Modify Date: 29/11/2010
'Crea el archivo de seguimientos.
'**************************************************************
On Error GoTo ErrHandler
  
Dim intFile As Integer

    intFile = FreeFile
    
    Open DatPath & "RECORDS.DAT" For Output As #intFile
        Print #intFile, "[INIT]"
        Print #intFile, "NumRecords=0"
    Close #intFile
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CreateRecordsFile de modUserRecords.bas")
End Sub

Public Sub CheckIsBeingFollowed(ByVal UserIndex As Integer)
'**************************************************************
'Author: Zama
'Last Modify Date: 17/01/2014
'Checks if is being followed and warns admins of its conection.
'**************************************************************
On Error GoTo ErrHandler
  
    
    Dim UserName As String
    UserName = UCase$(UserList(UserIndex).Name)
    
    Dim lCounter As Long
    For lCounter = 1 To NumRecords
        If Records(lCounter).Usuario = UserName Then
            Call SendData(SendTarget.ToAdminsButCounselorsAndRms, 0, PrepareMessageConsoleMsg("SEGUIMIENTO> " & UserName & " se conectó con ip: " & UserList(UserIndex).IP, FontTypeNames.FONTTYPE_FIGHT))
            Exit Sub
        End If
    Next lCounter
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CheckIsBeingFollowed de modUserRecords.bas")
End Sub
