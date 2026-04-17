Attribute VB_Name = "modHelperFunctions"
Option Explicit


''' This function returns true if the key pressed is a valid key
Public Function IsNumericInputKeyPressValid(ByRef KeyAscii As Integer, Optional ByVal OverrideKeyAscii As Boolean = False) As Boolean
    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            If OverrideKeyAscii Then KeyAscii = 0
            IsNumericInputKeyPressValid = False
            Exit Function
        End If
    End If
    
    IsNumericInputKeyPressValid = True
    Exit Function
End Function


Public Function SecondsToTimeString(ByVal Seconds As Long) As String
    SecondsToTimeString = Format$((Seconds \ 3600), "00") & ":" & _
          Format$((Seconds Mod 3600) \ 60, "00") & ":" & _
          Format$((Seconds Mod 60), "00")
End Function

