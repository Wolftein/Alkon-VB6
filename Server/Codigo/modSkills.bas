Attribute VB_Name = "modSkills"
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
' Module author: D'Artagnan (14/04/2015)
'

Option Explicit

Public Sub AddAssignedSkills(ByVal nUserIndex As Integer, ByVal skill As eSkill, ByVal Value As Byte)
'******************************************
'Author: D'Artagnan
'Date: 14/04/2015
'Add the specified amount of assigned skills.
'******************************************
On Error GoTo ErrHandler
  
    If GetSkills(nUserIndex, skill) >= 100 Then Exit Sub
    
    With UserList(nUserIndex)
        .Stats.AssignedSkills(skill) = .Stats.AssignedSkills(skill) + value
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AddAssignedSkills de modSkills.bas")
End Sub

Public Sub AddNaturalSkills(ByVal nUserIndex As Integer, ByVal skill As eSkill, ByVal Value As Byte)
'******************************************
'Author: D'Artagnan
'Date: 14/04/2015
'Add the specified amount of natural skills.
'******************************************
On Error GoTo ErrHandler
  
    If GetSkills(nUserIndex, skill) >= 100 Then Exit Sub
    
    With UserList(nUserIndex)
        .Stats.NaturalSkills(skill) = .Stats.NaturalSkills(skill) + value
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AddNaturalSkills de modSkills.bas")
End Sub

Public Function GetAssignedSkills(ByVal nUserIndex As Integer, ByVal skill As eSkill) As Byte
'******************************************
'Author: D'Artagnan
'Date: 14/04/2015
'
'******************************************
On Error GoTo ErrHandler
  
    GetAssignedSkills = UserList(nUserIndex).Stats.AssignedSkills(skill)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetAssignedSkills de modSkills.bas")
End Function

Public Function GetNaturalSkills(ByVal nUserIndex As Integer, ByVal skill As eSkill) As Byte
'******************************************
'Author: D'Artagnan
'Date: 14/04/2015
'
'******************************************
On Error GoTo ErrHandler
  
    GetNaturalSkills = UserList(nUserIndex).Stats.NaturalSkills(skill)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetNaturalSkills de modSkills.bas")
End Function

Public Function GetSkills(ByVal nUserIndex As Integer, ByVal skill As eSkill) As Byte
'******************************************
'Author: D'Artagnan
'Date: 14/04/2015
'Get the amount of skills including both natural and assigned.
'******************************************
On Error GoTo ErrHandler
  
    GetSkills = GetAssignedSkills(nUserIndex, skill) + GetNaturalSkills(nUserIndex, skill)
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetSkills de modSkills.bas")
End Function

Public Function NaturalSkillsAvailable(ByVal nUserIndex As Integer, ByVal skill As eSkill) As Boolean
'******************************************
'Author: D'Artagnan
'Date: 14/04/2015
'Return True if the specified skill can be trained.
'False otherwise.
'******************************************
On Error GoTo ErrHandler
  
    With UserList(nUserIndex)
        NaturalSkillsAvailable = GetNaturalSkills(nUserIndex, skill) < (.Stats.ELV * 2)
    End With
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function NaturalSkillsAvailable de modSkills.bas")
End Function

Public Sub ZeroSkills(ByVal nUserIndex As Integer, ByVal skill As eSkill, _
                      Optional ByVal bNatural As Boolean = True, _
                      Optional ByVal bAssigned As Boolean = True)
'******************************************
'Author: D'Artagnan
'Date: 14/04/2015
'Reset the amount of the specified skills.
'******************************************
On Error GoTo ErrHandler
  
    With UserList(nUserIndex).Stats
        If bNatural Then .NaturalSkills(skill) = 0
        If bAssigned Then .AssignedSkills(skill) = 0
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ZeroSkills de modSkills.bas")
End Sub
