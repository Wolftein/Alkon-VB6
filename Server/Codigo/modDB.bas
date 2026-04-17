Attribute VB_Name = "modDB"
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
Option Explicit

' DB Connection
Private conn As ADODB.Connection

Public Function ConnectDB() As Boolean
'***************************************************
'Author: ZaMa
'Creation Date: 09/06/2012
'Last Modification: -
'Connects to DB
'***************************************************
On Error GoTo ErrHandler
    
    Dim ConectionString As String
    ConectionString = _
        "Driver={" & GetVar(IniPath & "Server.ini", "DATABASE", "Driver") & "};" & _
        "Server=" & GetVar(IniPath & "Server.ini", "DATABASE", "Server") & ";" & _
        "Port=" & GetVar(IniPath & "Server.ini", "DATABASE", "Port") & ";" & _
        "Database=" & GetVar(IniPath & "Server.ini", "DATABASE", "Database") & ";" & _
        "UID=" & GetVar(IniPath & "Server.ini", "DATABASE", "UID") & ";" & _
        "Password=" & GetVar(IniPath & "Server.ini", "DATABASE", "Password") & ";"
        
    Set conn = New ADODB.Connection
    
    conn.CursorLocation = adUseClient
    
    conn.Open ConectionString
    
    ConnectDB = True
    
    Exit Function
ErrHandler:
    MsgBox "Imposible conectar a la DB: " & vbCrLf & _
        Err.Description
End Function

Public Function EscapeString(ByRef parameter As String)
On Error GoTo ErrHandler
  
    Dim ret As String

    EscapeString = Replace$(parameter, "'", "''")
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function EscapeString de modDB.bas")
End Function

Public Function closeDB() As Boolean
On Error GoTo ErrHandler
    
    Call conn.Close
    
    Set conn = Nothing

    closeDB = False
    

    Exit Function
ErrHandler:

End Function


Public Function ExecuteSql(ByRef Sql As String, Optional ByRef affectedRecords As Long = 0) As ADODB.Recordset
'***************************************************
'Author: ZaMa
'Creation Date: 09/06/2012
'Last Modification: -
'Executes a given Sql and returns a recordset.
'***************************************************
    'Dim nfile As Integer
    'nfile = FreeFile ' obtenemos un canal
    'Open App.Path & "\logs\queries.log" For Append Shared As #nfile
    'Print #nfile, Date & " " & Time & " " & Sql
    'Close #nfile
On Error GoTo ErrHandler
  
    'conn.CursorLocation = adUseClient
    Set ExecuteSql = conn.Execute(Sql, affectedRecords)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") Query: " & Sql & " en Function ExecuteSql de modDB.bas")
End Function

Public Function ExecuteSqlCommand(ByRef Cmd As ADODB.Command) As ADODB.Recordset
On Error GoTo ErrHandler
  
    Cmd.ActiveConnection = conn
    Set ExecuteSqlCommand = Cmd.Execute
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ExecuteSqlCommand de modDB.bas")
End Function

Public Sub ExecuteSQLScript(ByRef sFileName As String)
'***************************************************
'Author: D'Artagnan
'Last Modification: 28/03/2015
'Run the (restricted) SQL code contained in the specified file.
'***************************************************
On Error GoTo ErrHandler
  
    Dim f As Integer
    Dim I As Long
    Dim J As Long
    Dim sDelimiter As String
    Dim sNewDelimiter As String
    Dim sLine As String
    Dim sQueries() As String
    
    sDelimiter = ";"  ' Default SQL delimiter.
    sNewDelimiter = vbNullString
    I = -1
    f = FreeFile
    
    Open sFileName For Input As #f
    Do Until EOF(f)
        Line Input #f, sLine
        sLine = Trim$(sLine)
        
        ' Avoid empty lines.
        If LenB(sLine) > 0 Then
            ' Ignore inline comments.
            If Left$(sLine, 2) <> "--" Then
                If I = -1 Then
                    ReDim sQueries(0 To 0)
                    sQueries(0) = sLine
                    I = 0
                Else
                    ' Change delimiter if necessary.
                    If Left$(sLine, 9) = "DELIMITER" Then
                        sNewDelimiter = Trim$(mid$(sLine, 11))
                    Else
                        If Right$(sQueries(I), Len(sDelimiter)) = sDelimiter Then
                            I = I + 1
                            ReDim Preserve sQueries(0 To I)
                            sQueries(I) = sLine
                        Else
                            ' Assign to the same query until a delimiter is found.
                            sQueries(I) = sQueries(I) & " " & sLine
                        End If
                        
                        If LenB(sNewDelimiter) > 0 Then
                            sDelimiter = sNewDelimiter
                            sNewDelimiter = vbNullString
                        End If
                    End If
                End If
            End If
        End If
    Loop
    Close #f
    
    For J = 0 To I
        Call ExecuteSql(sQueries(J))
    Next J
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ExecuteSQLScript de modDB.bas")
End Sub


Public Function GetStringOr(ByRef RsField As ADODB.Field, ByRef DefaultValue As String)
    If Not IsNull(RsField) Then
        GetStringOr = CStr(RsField)
    Else
        GetStringOr = DefaultValue
    End If
End Function
