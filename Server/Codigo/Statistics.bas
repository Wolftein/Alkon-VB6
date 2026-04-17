Attribute VB_Name = "Statistics"
'**************************************************************
' modStatistics.bas - Takes statistics on the game for later study.
'
' Implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
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

Private Type fragLvlRace
    matrix(1 To 50, 1 To 5) As Long
End Type

Private Type fragLvlLvl
    matrix(1 To 50, 1 To 50) As Long
End Type

Private fragLvlRaceData(1 To 7) As fragLvlRace
Private fragLvlLvlData(1 To 7) As fragLvlLvl
Private fragAlignmentLvlData(1 To 50, 1 To 4) As Long

'Currency just in case.... chats are way TOO often...
Private keyOcurrencies(255) As Currency

Public Sub UserDisconnected(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    With UserList(UserIndex)
        'Update trainning time
        With .trainningData
            .trainningTime = .trainningTime + (GetTickCount() - .startTick) / 1000
            
            .startTick = GetTickCount()
        End With

    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UserDisconnected de Statistics.bas")
End Sub

Public Sub UserLevelUp(ByVal UserIndex As Integer, Optional ByVal LogMasteryPoint As Boolean = False)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler

    Dim handle As Integer
    
    With UserList(UserIndex)
        
        Dim Descrip As String
        
        If Not LogMasteryPoint Then
            Descrip = "Completó el nivel " & CStr(.Stats.ELV) & " en " & _
                CStr(.trainningData.trainningTime + (GetTickCount() - .trainningData.startTick) / 1000) & " segundos."
        Else
            Descrip = "Ganó una maestría. Ahora tiene: " & CStr(.Stats.MasteryPoints) & " en " & _
                CStr(.trainningData.trainningTime + (GetTickCount() - .trainningData.startTick) / 1000) & " segundos."
        End If
        
        Call SaveStatictisDB(.ID, Descrip)

        'Reset data
        .trainningData.trainningTime = 0
        .trainningData.startTick = GetTickCount()
    End With
    
    Exit Sub
ErrHandler:
    Call LogError("Error en UserLevelUp. Error: " & Err.Description & ". User: " & UserList(UserIndex).Name)
    If handle <> 0 Then Close handle
End Sub

Public Sub StoreFrag(ByVal killer As Integer, ByVal victim As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim clase As Integer
    Dim raza As Integer
    Dim Alignment As Integer
    
    If UserList(victim).Stats.ELV > ConstantesBalance.MaxLvl Or UserList(killer).Stats.ELV > ConstantesBalance.MaxLvl Then Exit Sub
    
    Select Case UserList(killer).clase
        Case eClass.Assasin
            clase = 1
        
        Case eClass.Bard
            clase = 2
        
        Case eClass.Mage
            clase = 3
        
        Case eClass.Paladin
            clase = 4
        
        Case eClass.Warrior
            clase = 5
        
        Case eClass.Cleric
            clase = 6
        
        Case Else
            Exit Sub
    End Select
    
    Select Case UserList(killer).raza
        Case eRaza.Elfo
            raza = 1
        
        Case eRaza.Drow
            raza = 2
        
        Case eRaza.Enano
            raza = 3
        
        Case eRaza.Gnomo
            raza = 4
        
        Case eRaza.Humano
            raza = 5
        
        Case Else
            Exit Sub
    End Select
    
    Alignment = UserList(killer).Faccion.Alignment
    
    fragLvlRaceData(clase).matrix(UserList(killer).Stats.ELV, raza) = fragLvlRaceData(clase).matrix(UserList(killer).Stats.ELV, raza) + 1
    
    fragLvlLvlData(clase).matrix(UserList(killer).Stats.ELV, UserList(victim).Stats.ELV) = fragLvlLvlData(clase).matrix(UserList(killer).Stats.ELV, UserList(victim).Stats.ELV) + 1
    
    fragAlignmentLvlData(UserList(killer).Stats.ELV, Alignment) = fragAlignmentLvlData(UserList(killer).Stats.ELV, Alignment) + 1
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub StoreFrag de Statistics.bas")
End Sub

Public Sub DumpStatistics()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim handle As Integer
    handle = FreeFile()
    
    Dim line As String
    Dim I As Long
    Dim J As Long
    
    Open ServerConfiguration.LogsPaths.GeneralPath & "frags.txt" For Output As handle
    
    'Save lvl vs lvl frag matrix for each class - we use GNU Octave's ASCII file format
    
    Print #handle, "# name: fragLvlLvl_Ase"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 50"
    Print #handle, "# columns: 50"
    
    For J = 1 To 50
        For I = 1 To 50
            line = line & " " & CStr(fragLvlLvlData(1).matrix(I, J))
        Next I
        
        Print #handle, line
        line = vbNullString
    Next J
    
    Print #handle, "# name: fragLvlLvl_Bar"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 50"
    Print #handle, "# columns: 50"
    
    For J = 1 To 50
        For I = 1 To 50
            line = line & " " & CStr(fragLvlLvlData(2).matrix(I, J))
        Next I
        
        Print #handle, line
        line = vbNullString
    Next J
    
    Print #handle, "# name: fragLvlLvl_Mag"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 50"
    Print #handle, "# columns: 50"
    
    For J = 1 To 50
        For I = 1 To 50
            line = line & " " & CStr(fragLvlLvlData(3).matrix(I, J))
        Next I
        
        Print #handle, line
        line = vbNullString
    Next J
    
    Print #handle, "# name: fragLvlLvl_Pal"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 50"
    Print #handle, "# columns: 50"
    
    For J = 1 To 50
        For I = 1 To 50
            line = line & " " & CStr(fragLvlLvlData(4).matrix(I, J))
        Next I
        
        Print #handle, line
        line = vbNullString
    Next J
    
    Print #handle, "# name: fragLvlLvl_Gue"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 50"
    Print #handle, "# columns: 50"
    
    For J = 1 To 50
        For I = 1 To 50
            line = line & " " & CStr(fragLvlLvlData(5).matrix(I, J))
        Next I
        
        Print #handle, line
        line = vbNullString
    Next J
    
    Print #handle, "# name: fragLvlLvl_Cle"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 50"
    Print #handle, "# columns: 50"
    
    For J = 1 To 50
        For I = 1 To 50
            line = line & " " & CStr(fragLvlLvlData(6).matrix(I, J))
        Next I
        
        Print #handle, line
        line = vbNullString
    Next J
    
    Print #handle, "# name: fragLvlLvl_Caz"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 50"
    Print #handle, "# columns: 50"
    
    For J = 1 To 50
        For I = 1 To 50
            line = line & " " & CStr(fragLvlLvlData(7).matrix(I, J))
        Next I
        
        Print #handle, line
        line = vbNullString
    Next J
    
    
    
    
    
    'Save lvl vs race frag matrix for each class - we use GNU Octave's ASCII file format
    
    Print #handle, "# name: fragLvlRace_Ase"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 5"
    Print #handle, "# columns: 50"
    
    For J = 1 To 5
        For I = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(1).matrix(I, J))
        Next I
        
        Print #handle, line
        line = vbNullString
    Next J
    
    Print #handle, "# name: fragLvlRace_Bar"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 5"
    Print #handle, "# columns: 50"
    
    For J = 1 To 5
        For I = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(2).matrix(I, J))
        Next I
        
        Print #handle, line
        line = vbNullString
    Next J
    
    Print #handle, "# name: fragLvlRace_Mag"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 5"
    Print #handle, "# columns: 50"
    
    For J = 1 To 5
        For I = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(3).matrix(I, J))
        Next I
        
        Print #handle, line
        line = vbNullString
    Next J
    
    Print #handle, "# name: fragLvlRace_Pal"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 5"
    Print #handle, "# columns: 50"
    
    For J = 1 To 5
        For I = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(4).matrix(I, J))
        Next I
        
        Print #handle, line
        line = vbNullString
    Next J
    
    Print #handle, "# name: fragLvlRace_Gue"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 5"
    Print #handle, "# columns: 50"
    
    For J = 1 To 5
        For I = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(5).matrix(I, J))
        Next I
        
        Print #handle, line
        line = vbNullString
    Next J
    
    Print #handle, "# name: fragLvlRace_Cle"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 5"
    Print #handle, "# columns: 50"
    
    For J = 1 To 5
        For I = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(6).matrix(I, J))
        Next I
        
        Print #handle, line
        line = vbNullString
    Next J
    
    Print #handle, "# name: fragLvlRace_Caz"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 5"
    Print #handle, "# columns: 50"
    
    For J = 1 To 5
        For I = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(7).matrix(I, J))
        Next I
        
        Print #handle, line
        line = vbNullString
    Next J
    
    
    
    
    
    
    'Save lvl vs class frag matrix for each race - we use GNU Octave's ASCII file format
    
    Print #handle, "# name: fragLvlClass_Elf"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 7"
    Print #handle, "# columns: 50"
    
    For J = 1 To 7
        For I = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(J).matrix(I, 1))
        Next I
        
        Print #handle, line
        line = vbNullString
    Next J
    
    Print #handle, "# name: fragLvlClass_Dar"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 7"
    Print #handle, "# columns: 50"
    
    For J = 1 To 7
        For I = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(J).matrix(I, 2))
        Next I
        
        Print #handle, line
        line = vbNullString
    Next J
    
    Print #handle, "# name: fragLvlClass_Dwa"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 7"
    Print #handle, "# columns: 50"
    
    For J = 1 To 7
        For I = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(J).matrix(I, 3))
        Next I
        
        Print #handle, line
        line = vbNullString
    Next J
    
    Print #handle, "# name: fragLvlClass_Gno"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 7"
    Print #handle, "# columns: 50"
    
    For J = 1 To 7
        For I = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(J).matrix(I, 4))
        Next I
        
        Print #handle, line
        line = vbNullString
    Next J
    
    Print #handle, "# name: fragLvlClass_Hum"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 7"
    Print #handle, "# columns: 50"
    
    For J = 1 To 7
        For I = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(J).matrix(I, 5))
        Next I
        
        Print #handle, line
        line = vbNullString
    Next J
    
    
    
    
    'Save lvl vs alignment frag matrix for each race - we use GNU Octave's ASCII file format
    
    Print #handle, "# name: fragAlignmentLvl"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 4"
    Print #handle, "# columns: 50"
    
    For J = 1 To 4
        For I = 1 To 50
            line = line & " " & CStr(fragAlignmentLvlData(I, J))
        Next I
        
        Print #handle, line
        line = vbNullString
    Next J
    
    Close handle
    
    
    
    'Dump Chat statistics
    handle = FreeFile()
    
    Open ServerConfiguration.LogsPaths.GeneralPath & "huffman.log" For Output As handle
    
    Dim Total As Currency
    
    'Compute total characters
    For I = 0 To 255
        Total = Total + keyOcurrencies(I)
    Next I
    
    'Show each character's ocurrencies
    If Total <> 0 Then
        For I = 0 To 255
            Print #handle, CStr(I) & "    " & CStr(Round(keyOcurrencies(I) / Total, 8))
        Next I
    End If
    
    Print #handle, "TOTAL =    " & CStr(Total)
    
    Close handle
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DumpStatistics de Statistics.bas")
End Sub

Public Sub ParseChat(ByRef S As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler
  

    Dim I As Long
    Dim Key As Integer
    
    For I = 1 To Len(S)
        Key = Asc(mid$(S, I, 1))
        
        keyOcurrencies(Key) = keyOcurrencies(Key) + 1
    Next I
    
    'Add a NULL-terminated to consider that possibility too....
    keyOcurrencies(0) = keyOcurrencies(0) + 1
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ParseChat de Statistics.bas")
End Sub
