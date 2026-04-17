VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicalButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Graphical Button
'Designed and implemented by ZaMa
'zaxtor@gmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************
Option Explicit

Private WithEvents ButtonHolder As Image
Attribute ButtonHolder.VB_VarHelpID = -1

Private RollOverButtonImage As Picture  'img holder
Private NormalButtonImage As Picture 'img holder
Private ClickButtonImage As Picture  'img holder
Private DisabledButtonImage As Picture  'img holder
Private ExtraButtonImage As Picture 'img holder

Private SoundClick As Boolean

Private FormActual As Form ' Form owner

Private OverButton As Boolean 'Indicates wether the mouse is over the button or not
Private Disabled As Boolean 'Indicates wether it's enabled or not
'''
Private InternalEventHandler As clsButtonEventHandler
Private InternalIndex As Integer

Public Property Get Name() As String
    Name = ButtonHolder.Name
End Property

Public Property Get Index() As Integer
    Index = InternalIndex
End Property

Public Property Let Index(NewIndex As Integer)
    InternalIndex = NewIndex
End Property

Public Property Get EventHandler() As clsButtonEventHandler
    Set EventHandler = InternalEventHandler
End Property

Public Property Set EventHandler(eh As clsButtonEventHandler)
    Set InternalEventHandler = eh
End Property
'''
Public Sub Initialize(ByRef imgButton As Variant, ByVal imgNormalPath As String, _
            ByVal imgRollOverPath As String, ByVal imgClickPath As String, ByRef Formulario As Form, _
            Optional ByVal imgDisabledPath As String, Optional ByVal StartDisabled As Boolean = False, _
            Optional ByVal MakeClickSound As Boolean = True, Optional ByVal ShowHandPointer As Boolean = True, _
            Optional ByVal imgExtraPath As String)
'***************************************************
'Author: ZaMa
'Last Modify Date: 22/12/09
'Sets the reference to the image container
'***************************************************
On Error GoTo ErrHandler
  
    '''
    Set InternalEventHandler = Nothing
    '''
    Call CleanForTermination
    
    Set ButtonHolder = imgButton
    Set FormActual = Formulario
    
    Set RollOverButtonImage = LoadPicture(imgRollOverPath)
    Set NormalButtonImage = LoadPicture(imgNormalPath)
    Set ClickButtonImage = LoadPicture(imgClickPath)
    Set DisabledButtonImage = LoadPicture(imgDisabledPath)
    Set ExtraButtonImage = LoadPicture(imgExtraPath)
    
    SoundClick = MakeClickSound
    
    With ButtonHolder
        If ShowHandPointer Then
            .MouseIcon = picMouseIcon
            .MousePointer = vbCustom
        End If
        
        If StartDisabled Then
            Disabled = True
            .Picture = DisabledButtonImage
        Else
            .Picture = NormalButtonImage
        End If
    End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Initialize de clsGraphicalButton.cls")
End Sub

Private Sub ButtonHolder_Click()
'***************************************************
'Author: ZaMa
'Last Modify Date: 02/01/2010
'Implements the mouse click event of the image button.
'Plays the Click Sound.
'***************************************************
On Error GoTo ErrHandler
  
    If Disabled Then Exit Sub
    '''
    If Not (InternalEventHandler Is Nothing) Then
        Call InternalEventHandler.OnGraphicalButtonClick(Me)
    End If

    If Not SoundClick Then Exit Sub
    
    Call Engine_Audio.PlayInterface(SND_CLICK)
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ButtonHolder_Click de clsGraphicalButton.cls")
End Sub

Private Sub ButtonHolder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: ZaMa
'Last Modify Date: 29/12/09
'Implements the mouse down event of the image button.
'Toggle to Click image button.
'***************************************************
On Error GoTo ErrHandler
  
    If Disabled Then Exit Sub
    
    ButtonHolder.Picture = ClickButtonImage
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ButtonHolder_MouseDown de clsGraphicalButton.cls")
End Sub

Private Sub ButtonHolder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: ZaMa
'Last Modify Date: 22/12/09
'Implements the mouse move event of the image button.
'Toggle to pressed image button.
'***************************************************
On Error GoTo ErrHandler
  
    If Disabled Then Exit Sub
    
    If Not OverButton Then
    
        'This helps when two buttons are close
        FormActual.LastButtonPressed.ToggleToNormal
        
        ButtonHolder.Picture = RollOverButtonImage
        OverButton = True
        
        Set FormActual.LastButtonPressed = Me
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ButtonHolder_MouseMove de clsGraphicalButton.cls")
End Sub

Public Sub ToggleToNormal()
'***************************************************
'Author: ZaMa
'Last Modify Date: 22/12/09
'Toggle to normal image button.
'***************************************************
On Error GoTo ErrHandler
  
    If Disabled Then Exit Sub
    
    If Not OverButton Then Exit Sub
    
    ButtonHolder.Picture = NormalButtonImage
    OverButton = False
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ToggleToNormal de clsGraphicalButton.cls")
End Sub

Public Sub ToggleExtra()
'***************************************************
'Author: Anagrama
'Last Modify Date: 08/01/2016
'Toggle to extra/default image button.
'***************************************************
On Error GoTo ErrHandler
    Dim tmpPic As Picture
    
    If ButtonHolder.Picture = NormalButtonImage Then
        Set tmpPic = NormalButtonImage
        Set NormalButtonImage = ExtraButtonImage
        ButtonHolder.Picture = NormalButtonImage
        Set ExtraButtonImage = tmpPic
    Else
        Set tmpPic = ExtraButtonImage
        Set ExtraButtonImage = NormalButtonImage
        ButtonHolder.Picture = NormalButtonImage
        Set NormalButtonImage = tmpPic
    End If

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ToggleExtra de clsGraphicalButton.cls")
End Sub

Public Sub SoundEnabled(ByVal Enabled As Boolean)
On Error GoTo ErrHandler
  
    SoundClick = Enabled
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SoundEnabled de clsGraphicalButton.cls")
End Sub


Private Sub ButtonHolder_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: ZaMa
'Last Modify Date: 29/12/09
'Implements the mouse up event of the image button.
'Toggle to RollOver image button.
'***************************************************
On Error GoTo ErrHandler
  
    If Disabled Then Exit Sub
    
    ButtonHolder.Picture = RollOverButtonImage
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ButtonHolder_MouseUp de clsGraphicalButton.cls")
End Sub

Public Sub EnableButton(Optional ByVal ToogleEnable As Boolean = True)
On Error GoTo ErrHandler
  
    If Disabled = Not ToogleEnable Then Exit Sub
    
    OverButton = False
    Disabled = Not ToogleEnable
    If Disabled Then
        ButtonHolder.Picture = DisabledButtonImage
    Else
        ButtonHolder.Picture = NormalButtonImage
    End If
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub EnableButton de clsGraphicalButton.cls")
End Sub

Public Function IsEnabled() As Boolean
On Error GoTo ErrHandler
  
    IsEnabled = Not Disabled
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function IsEnabled de clsGraphicalButton.cls")
End Function

Private Sub CleanForTermination()
    Set FormActual = Nothing

    Set ButtonHolder = Nothing
    
    Set RollOverButtonImage = Nothing  'img holder
    Set NormalButtonImage = Nothing  'img holder
    Set ClickButtonImage = Nothing   'img holder
    Set DisabledButtonImage = Nothing   'img holder
    Set ExtraButtonImage = Nothing  'img holder
End Sub

Private Sub Class_Terminate()
    Call CleanForTermination
End Sub
