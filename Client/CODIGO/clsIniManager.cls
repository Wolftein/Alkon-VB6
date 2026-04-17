VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsIniManager.cls - Loads INI files into memory and applies Binary Search to get values at high speed.
' Use it instead of GetVar when reading several values form the same file at once, otherwise it's not usefull.
' Based in the idea of AlejoLP and his clsLeerInis.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
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

''
'Loads a complete INI file into memory and sorts it's data and keys for faster searches.
'It is MUCH faster than GetPrivateProfileStringA if you search for several values within a file, otherwise stick to the API.
'It's particularly usefull when you can keep a whole file in memory such as NPCs.dat'
' Based in the idea of AlejoLP and his clsLeerInis.
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.1.0
' @date 20060501

'01/05/2006 - Juan Martín Sotuyo Dodero (Maraxus) - (juansotuyo@gmail.com)
'   - First Release
'
'01/04/2008 - Juan Martín Sotuyo Dodero (Maraxus) - (juansotuyo@gmail.com)
'   - Add: KeyExists method allows to check for valid section keys.
'
'10/10/2010 - Torres Patricio(Pato)
'   - Add:  AddNode method to add a main node.
'           AddKey method to add key into the refered main node.
'           DumpFile method to dump the data into a file.
'30/12/2017 - Torres Patricio(Pato)
'   - Add:  SaveOpenFile to dump the file using the same path used to initialize the class.
'           GetFileNameFromPath method to get a filename based on a full path.


Option Explicit
Option Base 0


''
'Structure that contains a value and it's key in a INI file
'
' @param    key String containing the key associated to the value.
' @param    value String containing the value of the INI entry.
' @see      MainNode
'

Private Type childNode
    Key As String
    value As String
End Type

''
'Structure that contains all info under a tag in a INI file.
'Such tags are indicated with the "[" and "]" characters.
'
' @param    name String containing the text within the "[" and "]" characters.
'It's the key used when searching for a main section of the INI data.
' @param    values Array of ChildNodes, each containing a value entry along with it's key.
' @param    numValues Number of entrys in the main node.

Private Type MainNode
    Name As String
    values() As childNode
    numValues As Integer
End Type

''
'Containts all Main sections of the loaded INI file
Private fileData() As MainNode

''
'Stores the total number of main sections in the loaded INI file
Private MainNodes As Long

''
'Stores the filename of the opened file
Private fileName As String

''
'Stores the full path of the opened file
Private filePath As String

Private Const BOOL_MAX_VALUE As Byte = 1
Private Const BYTE_MAX_VALUE As Byte = 255
Private Const INT_MAX_VALUE As Integer = 32767
Private Const LNG_MAX_VALUE As Long = 2147483647

''
'Default constructor. Does nothing.

Private Sub Class_Initialize()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/01/2006
'
'**************************************************************
  
End Sub

''
'Destroy every array and deallocates al memory.
'

Private Sub Class_Terminate()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/01/2006
'
'**************************************************************
On Error GoTo ErrHandler
  
    Dim I As Long
    
    'Clean up
    If MainNodes Then
        For I = 1 To MainNodes - 1
            Erase fileData(I).values
        Next I
        
        Erase fileData
    End If
    
    MainNodes = 0
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Class_Terminate de clsIniManager.cls")
End Sub

''
'Loads a INI file so it's values can be read. Must be called before being able to use GetValue.
'
' @param    file Complete path of the INI file to be loaded.
' @see      GetValue

Public Sub Initialize(ByVal File As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 27/07/2006
'Opens the requested file and loads it's data into memory
'**************************************************************
On Error GoTo ErrHandler
  
    Dim handle As Integer
    Dim text As String
    Dim Pos As Long
    
    'Prevent memory losses if we are attempting to reload a file....
    Call Class_Terminate
    
    'Get a free handle and start reading line by line until the end
    handle = FreeFile
    
    Open File For Input As handle
    
    filePath = File
    fileName = GetFileNameFromPath(filePath)
    
    Do Until EOF(handle)
        Line Input #handle, text
        
        'Is it null??
        If Len(text) Then
            'If it starts with '[' it is a main node or nothing (GetPrivateProfileStringA works this way), otherwise it's a value
            If Left$(text, 1) = "[" Then
                'If it has an ending ']' it's a main node, otherwise it's nothing
                Pos = InStr(2, text, "]")
                If Pos Then
                    'Add a main node
                    ReDim Preserve fileData(MainNodes) As MainNode
                    
                    fileData(MainNodes).Name = UCase$(Trim$(mid$(text, 2, Pos - 2)))
                    
                    MainNodes = MainNodes + 1
                End If
            Else
                'So it's a value. Check if it has a '=', otherwise it's nothing
                Pos = InStr(2, text, "=")
                If Pos Then
                    'Is it under any main node??
                    If MainNodes Then
                        With fileData(MainNodes - 1)
                            'Add it to the main node's value
                            ReDim Preserve .values(.numValues) As childNode
                            
                            .values(.numValues).value = Right$(text, Len(text) - Pos)
                            .values(.numValues).Key = UCase$(Left$(text, Pos - 1))
                            
                            .numValues = .numValues + 1
                        End With
                    End If
                End If
            End If
        End If
    Loop
    
    Close handle
    
    Dim I As Long
    
    If MainNodes Then
        'Sort main nodes to allow binary search
        Call SortMainNodes(0, MainNodes - 1)
        
        'Sort values of each node to allow binary search
        For I = 0 To MainNodes - 1
            If fileData(I).numValues Then _
                Call SortChildNodes(fileData(I), 0, fileData(I).numValues - 1)
        Next I
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Initialize de clsIniManager.cls")
End Sub

''
'Sorts all child nodes within the given MainNode alphabetically by their keys. Uses quicksort.
'
' @param    Node The MainNode whose values are to be sorted.
' @param    first The first index to consider when sorting.
' @param    last The last index to be considered when sorting.

Private Sub SortChildNodes(ByRef Node As MainNode, ByVal First As Integer, ByVal Last As Integer)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/01/2006
'Sorts the list of values in a given MainNode using quicksort,
'this allows the use of Binary Search for faster searches
'**************************************************************
    Dim Min As Integer      'First item in the list
    Dim Max As Integer      'Last item in the list
    Dim comp As String      'Item used to compare
On Error GoTo ErrHandler
  
    Dim Temp As childNode
    
    Min = First
    Max = Last
    
    With Node
        comp = .values((Min + Max) \ 2).Key
    
        Do While Min <= Max
            Do While .values(Min).Key < comp And Min < Last
                Min = Min + 1
            Loop
            Do While .values(Max).Key > comp And Max > First
                Max = Max - 1
            Loop
            If Min <= Max Then
                Temp = .values(Min)
                .values(Min) = .values(Max)
                .values(Max) = Temp
                Min = Min + 1
                Max = Max - 1
            End If
        Loop
    End With
    
    If First < Max Then SortChildNodes Node, First, Max
    If Min < Last Then SortChildNodes Node, Min, Last
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SortChildNodes de clsIniManager.cls")
End Sub

''
'Sorts all main nodes in the loaded INI file alphabetically by their names. Uses quicksort.
'
' @param    first The first index to consider when sorting.
' @param    last The last index to be considered when sorting.

Private Sub SortMainNodes(ByVal First As Integer, ByVal Last As Integer)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/01/2006
'Sorts the MainNodes list using quicksort,
'this allows the use of Binary Search for faster searches
'**************************************************************
    Dim Min As Integer      'First item in the list
    Dim Max As Integer      'Last item in the list
    Dim comp As String      'Item used to compare
On Error GoTo ErrHandler
  
    Dim Temp As MainNode
    
    Min = First
    Max = Last
    
    comp = fileData((Min + Max) \ 2).Name
    
    Do While Min <= Max
        Do While fileData(Min).Name < comp And Min < Last
            Min = Min + 1
        Loop
        Do While fileData(Max).Name > comp And Max > First
            Max = Max - 1
        Loop
        If Min <= Max Then
            Temp = fileData(Min)
            fileData(Min) = fileData(Max)
            fileData(Max) = Temp
            Min = Min + 1
            Max = Max - 1
        End If
    Loop
    
    If First < Max Then SortMainNodes First, Max
    If Min < Last Then SortMainNodes Min, Last
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SortMainNodes de clsIniManager.cls")
End Sub

''
'Searches for a given key within a given main section and if it exists retrieves it's value, otherwise a null string
'
' @param    Main The name of the main section in which we will be searching.
' @param    key The key of the value we are looking for.
' @returns  The value asociated with the given key under the requeted main section of the INI file or a null string if it's not found.

Public Function GetValue(ByVal Main As String, ByVal Key As String) As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/01/2006
'Returns a value if the key and main node exist, or a nullstring otherwise
'**************************************************************
On Error GoTo ErrHandler
  
    Dim I As Long
    Dim J As Long
    
    'Search for the main node
    I = FindMain(UCase$(Main))
    
    If I >= 0 Then
        'If valid, binary search among keys
        J = FindKey(fileData(I), UCase$(Key))
        
        'If we found it we return it
        If J >= 0 Then GetValue = fileData(I).values(J).value
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetValue de clsIniManager.cls")
End Function

''
'Searches for a given key within a given main section and if it exists retrieves it's value, otherwise a null string
'
' @param    Main The name of the main section in which we will be searching.
' @param    key The key of the value we are looking for.
' @returns  The value asociated with the given key under the requeted main section of the INI file or a null string if it's not found.

Public Sub ChangeValue(ByVal Main As String, ByVal Key As String, ByVal value As String)
'**************************************************************
'Author: ZaMa
'Last Modify Date: 27/05/2009
'If the key and main node exist, changes the value
'**************************************************************
On Error GoTo ErrHandler
  
    Dim I As Long
    Dim J As Long
    
    'Search for the main node
    I = FindMain(UCase$(Main))
    
    If I < 0 Then I = AddNode(Main)
    
    'If valid, binary search among keys
    J = FindKey(fileData(I), UCase$(Key))
    
    If J < 0 Then
        Call AddKey(Main, Key, value)
    Else
        'If we found it we change it
        fileData(I).values(J).value = value
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ChangeValue de clsIniManager.cls")
End Sub

''
'Searches for a given key within a given main node and returns the index in which it's stored or the negation of the index in which it should be if not found.
'
' @param    Node The MainNode among whose value entries we will be searching.
' @param    key The key of the value we are looking for.
' @returns  The index in which the value with the key we are looking for is stored or the negation of the index in which it should be if not found.

Private Function FindKey(ByRef Node As MainNode, ByVal Key As String) As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/01/2006
'Returns the index of the value which key matches the requested one,
'or the negation of the position were it should be if not found
'**************************************************************
On Error GoTo ErrHandler
  
    Dim Min As Long
    Dim Max As Long
    Dim mid As Long
    
    Min = 0
    Max = Node.numValues - 1
    
    Do While Min <= Max
        mid = (Min + Max) \ 2
        
        If Node.values(mid).Key < Key Then
            Min = mid + 1
        ElseIf Node.values(mid).Key > Key Then
            Max = mid - 1
        Else
            'We found it
            FindKey = mid
            Exit Function
        End If
    Loop
    
    'Not found, return the negation of the position where it should be
    '(all higher values are to the right of the list and lower values are to the left)
    FindKey = Not Min
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function FindKey de clsIniManager.cls")
End Function

''
'Searches for a main section with the given name within the loaded INI file and returns the index in which it's stored or the negation of the index in which it should be if not found.
'
' @param    name The name of the MainNode we are looking for.
' @returns  The index in which the main section we are looking for is stored or the negation of the index in which it should be if not found.

Private Function FindMain(ByVal Name As String) As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/01/2006
'Returns the index of the MainNode which name matches the requested one,
'or the negation of the position were it should be if not found
'**************************************************************
On Error GoTo ErrHandler
  
    Dim Min As Long
    Dim Max As Long
    Dim mid As Long
    
    Min = 0
    Max = MainNodes - 1
    
    Do While Min <= Max
        mid = (Min + Max) \ 2
        
        If fileData(mid).Name < Name Then
            Min = mid + 1
        ElseIf fileData(mid).Name > Name Then
            Max = mid - 1
        Else
            'We found it
            FindMain = mid
            Exit Function
        End If
    Loop
    
    'Not found, return the negation of the position where it should be
    '(all higher values are to the right of the list and lower values are to the left)
    FindMain = Not Min
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function FindMain de clsIniManager.cls")
End Function

''
'Checks wether a given key exists or not.
'
' @param    name    The name of the element whose existance is being checked.
' @returns  True if the key exists, false otherwise.

Public Function KeyExists(ByVal Name As String) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 04/01/2008
'Returns true of the key exists, false otherwise.
'**************************************************************
On Error GoTo ErrHandler
  
    KeyExists = FindMain(UCase$(Name)) >= 0
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function KeyExists de clsIniManager.cls")
End Function

Private Function AddNode(ByRef Name As String) As Long
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 10/10/2010
'
'**************************************************************
On Error GoTo ErrHandler
  
Dim I As Long
Dim MainPos As Long

Name = UCase$(Name)

'Add a main node
ReDim Preserve fileData(MainNodes) As MainNode

If MainNodes Then
    I = MainNodes - 1
    MainPos = Not FindMain(Name)
    
    Do While I >= MainPos
        fileData(I + 1) = fileData(I)
        I = I - 1
    Loop
    
    fileData(MainPos).numValues = 0
    Erase fileData(MainPos).values
End If

fileData(MainPos).Name = Name

MainNodes = MainNodes + 1

AddNode = MainPos
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function AddNode de clsIniManager.cls")
End Function

Private Function AddKey(ByRef Main As String, ByRef Key As String, Optional ByRef value As String = vbNullString) As Long
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 10/10/2010
'
'**************************************************************
On Error GoTo ErrHandler
  
Dim MainPos As Long
Dim KeyPos As Long
Dim I As Long

Main = UCase$(Main)
Key = UCase$(Key)

MainPos = FindMain(Main)

If MainPos < 0 Then MainPos = AddNode(Main)

With fileData(MainPos)
    'Add it to the main node's value
    ReDim Preserve .values(.numValues) As childNode
    
    If .numValues > 0 Then
        I = .numValues - 1
        KeyPos = Not FindKey(fileData(MainPos), Key)
        
        Do While I >= KeyPos
            .values(I + 1) = .values(I)
            I = I - 1
        Loop
    End If
    
    .values(KeyPos).Key = Key
    .values(KeyPos).value = value
    
    .numValues = .numValues + 1
End With

AddKey = KeyPos
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function AddKey de clsIniManager.cls")
End Function

Public Sub SaveOpenFile()
    '**************************************************************
    'Author: Alejandro Masolini (Nightw)
    'Last Modify Date: 30/12/2017
    ' Saves the file based on the path used to initialize the class
    '**************************************************************
On Error GoTo ErrHandler
    Call DumpFile(filePath)
    
    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SaveOpenFile de clsIniManager.cls")
End Sub

Public Sub DumpFile(ByRef File As String)
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 10/10/2010
'
'**************************************************************
On Error GoTo ErrHandler
  
Dim hFile As Integer
Dim I As Long
Dim J As Long

hFile = FreeFile()

Open File For Output As hFile

    For I = 0 To MainNodes - 1
        With fileData(I)
            Print #hFile, "[" & .Name & "]"
            
            For J = 0 To .numValues - 1
                Print #hFile, .values(J).Key & "=" & .values(J).value
            Next J
            
            Print #hFile, ""
        End With
    Next I
    
Close hFile
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub DumpFile de clsIniManager.cls")
End Sub


''
'Searches for a given integer key within a given main section and if it exists retrieves it's value, otherwise a default value
'
' @param    Main The name of the main section in which we will be searching.
' @param    key The key of the value we are looking for.
' @param    defaultValue A default value if the read key was outside the boundaries or not found
' @returns  The value asociated with the given key under the requeted main section of the INI file or a null string if it's not found.

Public Function GetValueInt(ByRef Main As String, ByRef Key As String, Optional ByVal defaultValue As Integer = 0) As Integer
'**************************************************************
'Author: Alejandro Masolini
'Last Modify Date: 28/12/2017
'Returns an integer value if the key and main node exist, or a default value if it doesnt
'**************************************************************
On Error GoTo ErrHandler
  
    Dim tmpValue As String
    Dim tmpValueNumeric As Variant
    tmpValue = GetValue(Main, Key)
    
    ' Check if the value is a number
    If Not IsNumeric(tmpValue) Then
        GetValueInt = defaultValue
        Exit Function
    End If
    
    ' Check if the value is greater than the max expected value
    If Val(tmpValue) > INT_MAX_VALUE Then
        GetValueInt = defaultValue
        Exit Function
    End If
    
    GetValueInt = CInt(tmpValue)

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetValueInt de clsIniManager.cls")
End Function

''
'Searches for a given byte key within a given main section and if it exists retrieves it's value, otherwise a default value
'
' @param    Main The name of the main section in which we will be searching.
' @param    key The key of the value we are looking for.
' @param    defaultValue A default value if the read key was outside the boundaries or not found
' @returns  The value asociated with the given key under the requeted main section of the INI file or a null string if it's not found.

Public Function GetValueByte(ByRef Main As String, ByRef Key As String, Optional ByVal defaultValue As Byte = 0) As Byte
'**************************************************************
'Author: Alejandro Masolini
'Last Modify Date: 28/12/2017
'Returns an integer value if the key and main node exist, or a default value if it doesnt
'**************************************************************
On Error GoTo ErrHandler
  
    Dim tmpValue As String
    Dim tmpValueNumeric As Variant
    tmpValue = GetValue(Main, Key)
    
    ' Check if the value is a number
    If Not IsNumeric(tmpValue) Then
        GetValueByte = defaultValue
        Exit Function
    End If
    
    ' Check if the value is greater than the max expected value
    If Val(tmpValue) > BYTE_MAX_VALUE Then
        GetValueByte = defaultValue
        Exit Function
    End If
    
    GetValueByte = CByte(tmpValue)

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetValueByte de clsIniManager.cls")
End Function

''
'Searches for a given boolean key within a given main section and if it exists retrieves it's value, otherwise a default value
'
' @param    Main The name of the main section in which we will be searching.
' @param    key The key of the value we are looking for.
' @param    defaultValue A default value if the read key was outside the boundaries or not found
' @returns  The value asociated with the given key under the requeted main section of the INI file or a null string if it's not found.

Public Function GetValueBoolean(ByRef Main As String, ByRef Key As String, Optional ByVal defaultValue As Boolean = True) As Boolean
'**************************************************************
'Author: Alejandro Masolini
'Last Modify Date: 28/12/2017
'Returns an integer value if the key and main node exist, or a default value if it doesnt
'**************************************************************
On Error GoTo ErrHandler
  
    Dim tmpValue As String
    tmpValue = UCase(GetValue(Main, Key))
    
    ' Check if the value is a number
    If Not IsNumeric(tmpValue) Then
        ' Check if we got a representation of a boolean like "True" or "False"
        If tmpValue = "TRUE" Then
            GetValueBoolean = True
            Exit Function
        ElseIf tmpValue = "FALSE" Then
            GetValueBoolean = False
            Exit Function
        Else
            GetValueBoolean = defaultValue
            Exit Function
        End If
    End If
    
    ' Check if the value is greater than the max expected value
    If Val(tmpValue) > BOOL_MAX_VALUE Or Val(tmpValue) < -1 Then
        GetValueBoolean = defaultValue
        Exit Function
    End If
    
    GetValueBoolean = CBool(tmpValue)

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetValueByte de clsIniManager.cls")
End Function

''
'Searches for a given string key within a given main section and if it exists retrieves it's value, otherwise a null string
'
' @param    Main The name of the main section in which we will be searching.
' @param    key The key of the value we are looking for.
' @param    defaultValue A default value if the read key was outside the boundaries or not found
' @returns  The value asociated with the given key under the requeted main section of the INI file or a null string if it's not found.

Public Function GetValueString(ByRef Main As String, ByRef Key As String, Optional ByRef defaultValue As String = "") As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/01/2006
'Returns a value if the key and main node exist, or a nullstring otherwise
'**************************************************************
On Error GoTo ErrHandler
  
    Dim tmpValue As String
    
    tmpValue = GetValue(Main, Key)
    
    If tmpValue = vbNullString Then
        GetValueString = defaultValue
        Exit Function
    End If
    
    GetValueString = tmpValue
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetValue de clsIniManager.cls")
End Function


''
'Get a filename based on a full path.
'
' @param    strFullPath The full path
' @returns  The filename based on a full path
Private Function GetFileNameFromPath(strFullPath As String) As String
'**************************************************************
'Author: Alejandro Masolini
'Last Modify Date: 30/12/2017
'Returns a filename based on a full path
'**************************************************************
On Error GoTo ErrHandler

    GetFileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
    
    Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetFileNameFromPath de clsIniManager.cls")
End Function



