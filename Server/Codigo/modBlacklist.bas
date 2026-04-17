Attribute VB_Name = "modBlacklist"
'**************************************************************************
'Argentum Online
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
'**************************************************************************

'
' Module author: D'Artagnan (04/01/2015)
'

Option Explicit

Public Sub BlacklistAppend(ByRef sNickname As String)
'******************************************
'Author: D'Artagnan
'Date: 04/01/2015
'Insert the specified nickname into the blacklist table.
'******************************************
On Error GoTo ErrHandler
  
    If LenB(sNickname) = 0 Then Exit Sub
    
    
    
    Call ExecuteSql("INSERT INTO NICKNAMES_BLACKLIST VALUES ('" & Left$(UCase$(sNickname), 45) & "')")
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub BlacklistAppend de modBlacklist.bas")
End Sub

Public Function BlacklistGetNicknames() As String
'******************************************
'Author: D'Artagnan
'Date: 04/01/2015
'Return a string containing invalid nicknames.
'******************************************
On Error GoTo ErrHandler
  
    Dim rsQuery As ADODB.Recordset
    Dim sCurrentNickname As String
    
    Set rsQuery = ExecuteSql("SELECT NAME FROM NICKNAMES_BLACKLIST")
    
    While Not rsQuery.EOF
        sCurrentNickname = CStr(rsQuery.Fields(0))
        
        If LenB(BlacklistGetNicknames) = 0 Then
            BlacklistGetNicknames = sCurrentNickname
        Else
            BlacklistGetNicknames = BlacklistGetNicknames & ", " & sCurrentNickname
        End If
    Wend
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function BlacklistGetNicknames de modBlacklist.bas")
End Function

Public Function BlacklistIsValidNickname(ByRef sNickname As String) As Boolean
'******************************************
'Author: D'Artagnan
'Date: 04/01/2015
'
'******************************************
On Error GoTo ErrHandler
  
    Dim rsQuery As ADODB.Recordset
    
    ' Look for the specified nickname in the database table.
    Set rsQuery = ExecuteSql("SELECT COUNT(1) FROM NICKNAMES_BLACKLIST WHERE NAME='" & UCase$(sNickname) & "'")
    
    ' Retrieved count must be zero.
    BlacklistIsValidNickname = CInt(rsQuery.Fields(0)) = 0
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function BlacklistIsValidNickname de modBlacklist.bas")
End Function

Public Sub BlacklistRemove(ByRef sNickname As String)
'******************************************
'Author: D'Artagnan
'Date: 04/01/2015
'Delete the specified nickname from the blacklist table.
'******************************************
On Error GoTo ErrHandler
  
    If LenB(sNickname) = 0 Then Exit Sub
    
    Call ExecuteSql("DELETE FROM NICKNAMES_BLACKLIST WHERE NAME='" & UCase$(sNickname) & "'")
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub BlacklistRemove de modBlacklist.bas")
End Sub
