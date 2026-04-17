Attribute VB_Name = "modAccount"
' Programado por maTih.-
 
Option Explicit
 
'Máximos pjs por cuenta.
Public Const ACCPJS          As Byte = 8

Public Type tAccountConnectingMetadata
    UserRace As Byte
    UserClass As Byte
    UserGender As Byte
End Type
 
Public Type Acc_Chars
       Body                  As Integer
       Head                  As Integer
       Arma                  As Integer
       Escudo                As Integer
       Casco                 As Integer
End Type
 
Public Type tAccChars
    Char_Name           As String
    Char_Map_Name       As String
    Char_Nivel          As Byte
    Char_Muerto         As Boolean
    Char_Character      As Acc_Chars
    bSailing            As Boolean
    Alignment           As Byte
    IdGuild             As Long
    GuildName           As String
    JailRemainingTime   As Long
    Banned              As Boolean
End Type
 
Public Type tAccData
       Acc_Char_Selected     As Byte
       Acc_Name              As String
       Acc_Password          As String
       Acc_New_Password      As String
       Acc_Email             As String
       Acc_Pregunta          As String
       Acc_Recovering        As Boolean
       Acc_Respuesta         As String
       Acc_Char(1 To ACCPJS) As tAccChars
       Acc_Token             As String
       Acc_Waiting_Response  As Boolean
       Acc_Waiting_CharName  As String
       nCharCount            As Integer
End Type
 
Public Acc_Data              As tAccData
Public AccountConnecting     As tAccountConnectingMetadata

Public EMPTY_CHAR_DATA As tAccChars

Public mDevice(0 To 7)       As Long

Private Sub AccountCleanGUISlot(ByVal nCharSlot As Byte)
'******************************************
'Author: D'Artagnan
'Date: 19/07/2014
'Clean character picture and info.
'******************************************
On Error GoTo ErrHandler
  
    With frmAccount
        .picPJ(nCharSlot - 1).Cls
        .lblInfoChar(nCharSlot - 1).Caption = vbNullString
        .lblGuildName(nCharSlot - 1).Caption = vbNullString
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AccountCleanGUISlot de modAccount.bas")
End Sub

Public Function AccountGetToken() As String
'******************************************
'Author: D'Artagnan
'Date: 19/07/2014
'Return the current account token, if any.
'Otherwise, a null string is retrieved.
'******************************************
On Error GoTo ErrHandler
  
    If LenB(Acc_Data.Acc_Token) > 0 Then
        AccountGetToken = Acc_Data.Acc_Token
    Else
        ' Null character means there's no token available.
        AccountGetToken = Chr(0)
    End If
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function AccountGetToken de modAccount.bas")
End Function

Public Sub AccountRemoveCharacter(ByVal nCharSlot As Byte)
'******************************************
'Author: D'Artagnan
'Date: 19/07/2014
'Remove the character at the specified slot.
'Reorder remaining characters to fill the free slot.
'Allways check last slot data and flush char name(Luke)
'23/06/2018: IglorioN - Clear slot while moving backwards after removing char
'******************************************
On Error GoTo ErrHandler
  
    Dim I As Integer
    Dim tempCharacter As tAccChars
    Dim emptyCharacter As tAccChars
    
    Call AccountCleanGUISlot(nCharSlot)
    
    ' Reorder array if necessary.
    If nCharSlot < Acc_Data.nCharCount Then
        For I = nCharSlot + 1 To Acc_Data.nCharCount
            ' Get character data at the current slot.
            tempCharacter = Acc_Data.Acc_Char(I)
            
            If tempCharacter.Char_Name <> vbNullString Then
                ' Store it in the previous slot.
                Acc_Data.Acc_Char(I - 1) = tempCharacter
                ' Clear the slot that is moving
                Acc_Data.Acc_Char(I) = emptyCharacter
                Call AccountCleanGUISlot(I)
            End If
        Next I
    Else
        Acc_Data.Acc_Char(nCharSlot) = emptyCharacter
    End If
    
    ' Cleaning last slot char data if it still has
    If Acc_Data.Acc_Char(8).Char_Name <> vbNullString Then
        Acc_Data.Acc_Char(8).Char_Name = vbNullString
    End If
    
    Acc_Data.nCharCount = Acc_Data.nCharCount - 1
    
    ' Deselect current slot.
    If Acc_Data.Acc_Char_Selected = nCharSlot Then
        Acc_Data.Acc_Char_Selected = 0
        'frmAccount.cmdLog.Enabled = False
    End If
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AccountRemoveCharacter de modAccount.bas")
End Sub

Public Sub AccountReset()
'******************************************
'Author: D'Artagnan
'Date: 19/07/2014
'Reset account data.
'******************************************
On Error GoTo ErrHandler
  
    Dim I As Long
    Dim emptyData As tAccData
    Dim sToken As String
    
    sToken = Acc_Data.Acc_Token
    Acc_Data = emptyData
    Acc_Data.Acc_Token = sToken
    
    For I = 1 To ACCPJS
        Call AccountCleanGUISlot(I)
    Next I
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AccountReset de modAccount.bas")
End Sub

Public Sub Agregar_Personaje(ByVal char_Slot As Byte, ByRef acc_Type As tAccChars)
 
'
' @ Agrega un personaje a la cuenta.
On Error GoTo ErrHandler
  
 
With Acc_Data
     .Acc_Char(char_Slot) = acc_Type
     
     
     
     If (.Acc_Waiting_Response = True) Then
        If (.Acc_Waiting_CharName = .Acc_Char(char_Slot).Char_Name) Then
            Unload frmCrearPersonaje
            'frmAccount.SetFocus
            .Acc_Waiting_CharName = vbNullString
            .Acc_Waiting_Response = False
        End If
    End If
    
End With
 
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Agregar_Personaje de modAccount.bas")
End Sub

Public Sub Prepare_And_Connect(ByVal Mode As E_MODO, Optional ByRef sIPAddress As String = vbNullString, Optional ByVal nPort As Integer = 0)
On Error GoTo ErrHandler
  
    '
    ' @ Prepara y conecta el socket.
     
    Call frmConnect.CheckServers
    
    EstadoLogin = Mode
    If (Protocol.IsConnected()) Then
        Call HandleLogin
    Else
        Call Protocol.Connect(IIf(LenB(sIPAddress) > 0, sIPAddress, CurServerIp), IIf(nPort > 0, nPort, CurServerPort))
    End If
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Prepare_And_Connect de modAccount.bas")
End Sub
 
Public Sub Set_Acc_Data_To_Create()
 
'
' @ Setea los datos de la cuenta a crearse.
On Error GoTo ErrHandler
  
 
    With Acc_Data
     
        .Acc_Name = frmAccountCreate.txtName.text
        .Acc_Password = MD5.GetMD5String(frmAccountCreate.txtPassword.text)
    
        .Acc_Email = frmAccountCreate.txtEmail.text
        .Acc_Pregunta = "0000000000"
        .Acc_Respuesta = frmAccountCreate.txtSecurityCode.text
    End With

    Call MD5.MD5Reset
 
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Set_Acc_Data_To_Create de modAccount.bas")
End Sub
 
Public Sub Set_Acc_Data_To_Login()
 
'
' @ Setea los datos de la cuenta para conectarse.
On Error GoTo ErrHandler
  
    
With Acc_Data
     .Acc_Name = frmConnect.txtNombre.text
     .Acc_Password = MD5.GetMD5String(frmConnect.txtPasswd.text)
End With

MD5.MD5Reset
 
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Set_Acc_Data_To_Login de modAccount.bas")
End Sub
 
Public Sub Set_Acc_Data_To_Recover()
'
' @ Setea los datos de la cuenta para recuperar.
On Error GoTo ErrHandler
  
With Acc_Data
    .Acc_Name = frmAccountRecover.txtName.text
    .Acc_Token = MD5.GetMD5String(frmConnect.txtPasswd.text)
End With

MD5.MD5Reset
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Set_Acc_Data_To_Recover de modAccount.bas")
End Sub

Public Sub Set_Acc_Data_To_Delete()
'
' @ Set account token to validate char deletion on the server
On Error GoTo ErrHandler
  
With Acc_Data
    .Acc_Respuesta = frmDeleteCharValidation.txtToken.text
End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Set_Acc_Data_To_Delete de modAccount.bas")
End Sub

Public Function GetNickColorByAlignment(ByVal Alignment As Byte)
    Select Case Alignment
        Case eCharacterAlignment.Newbie
            GetNickColorByAlignment = RGB(242, 192, 41)
        Case eCharacterAlignment.Neutral
            GetNickColorByAlignment = RGB(155, 155, 155)
        Case eCharacterAlignment.FactionRoyal
            GetNickColorByAlignment = RGB(0, 128, 255)
        Case eCharacterAlignment.FactionLegion
            GetNickColorByAlignment = RGB(255, 0, 0)
    End Select

End Function

Public Sub Draw_Char_Slot(ByVal char_Slot As Byte, ByVal color As Long)
 
'
' @ Dibuja un personaje.
On Error GoTo ErrHandler
    Dim xOffset As Integer
    Dim yOffset As Integer
    
    xOffset = 24
    yOffset = 40
        
    With Acc_Data.Acc_Char(char_Slot)
        Dim t_Grh      As Grh
        Dim t_Position As Integer
                
        frmAccount.lblInfoChar(char_Slot - 1).Caption = .Char_Name & vbNewLine
        frmAccount.lblGuildName(char_Slot - 1).Caption = IIf(.GuildName <> vbNullString, "<" & .GuildName & ">", vbNullString)
        frmAccount.lblInfoChar(char_Slot - 1).ForeColor = GetNickColorByAlignment(.Alignment)
        frmAccount.lblGuildName(char_Slot - 1).ForeColor = frmAccount.lblInfoChar(char_Slot - 1).ForeColor

        'Char muerto?
        If (.Char_Muerto = True) Then
            t_Grh = BodyData(8).Walk(3)
                    
            Call DrawGrh(t_Grh, xOffset, yOffset, GetDepth(3, , , 2), 1, 0, color)
         
            t_Grh = HeadData(CASPER_HEAD).Head(3)
                        
            Call DrawGrh(t_Grh, xOffset + BodyData(8).HeadOffset.X, yOffset + BodyData(8).HeadOffset.Y, GetDepth(3, , , 1), 1, 0, color)

        Else
        
            Dim bSailing As Boolean
            Dim nHeadOffset As Integer
            
            bSailing = .bSailing
            
            'Está vivo, dibuja.
            With .Char_Character
                'Cuerpo
                If (.Body <> 0) And (.Body <= UBound(BodyData())) Then
                    t_Grh = BodyData(.Body).Walk(3)
                    
                    'DrawGrh t_Grh, IIf(bSailing, 8, 19 + xOffset), IIf(bSailing, 0, 15 + yOffset), char_Slot
                    
                    Call DrawGrh(t_Grh, IIf(bSailing, 8, xOffset), IIf(bSailing, -8, yOffset), GetDepth(3, , , 1), Not bSailing, 0, color)
                    
                End If
                
                '   Just body if sailing.
                If Not bSailing Then
                
                    'Cabeza
                    If (.Head <> 0) And (.Head <= UBound(HeadData())) Then
                        t_Grh = HeadData(.Head).Head(3)
                        
                        Call DrawGrh(t_Grh, xOffset + BodyData(.Body).HeadOffset.X, yOffset + BodyData(.Body).HeadOffset.Y, GetDepth(3, , , 2), 1, 0, color)
                    End If
                  
                    'Arma.
                    If (.Arma > 0 And .Arma <> 2) Then
                        t_Grh = WeaponAnimData(.Arma).WeaponWalk(3)
                        
                        Call DrawGrh(t_Grh, xOffset, yOffset, GetDepth(3, , , 6), 1, 0, color)
                    End If
                
                    'Casco
                    If (.Casco > 0 And .Casco <> 2) Then
                        t_Grh = CascoAnimData(.Casco).Head(3)
                        
                        Call DrawGrh(t_Grh, xOffset + BodyData(.Body).HeadOffset.X, yOffset + BodyData(.Body).HeadOffset.Y + OFFSET_HEAD, GetDepth(3, , , 4), 1, 0, color)
                    End If
                    
                    'Escudo
                    If (.Escudo > 0 And .Escudo <> 2) Then
                        t_Grh = ShieldAnimData(.Escudo).ShieldWalk(3)
                        
                       Call DrawGrh(t_Grh, xOffset, yOffset, GetDepth(3, , , 5), 1, 0, color)
                    End If
                End If
            End With
            
        End If
        
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Draw_Char_Slot de modAccount.bas")
End Sub


