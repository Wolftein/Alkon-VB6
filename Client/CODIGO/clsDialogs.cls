VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDialogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsDialogs.cls - Chats over head manager.
' Creates and manipulates chats over heads, rendering and updating them.
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
' Creates and manipulates chats over heads, rendering and updating them.
'
' @file     clsDialogs.cls
' @author   Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.0.0
' @date     20070729

Option Explicit
Option Base 0

''
' Maximum number of dialogs that can exist.
Private Const MAX_DIALOGS As Byte = 100

''
' Maximum length of any dialog line without having to split it.
Private Const MAX_LENGTH As Byte = 18

''
' Number of milliseconds to add to the lifetime per dialog character
Private Const MS_PER_CHAR As Byte = 100

''
' Number of extra milliseconds to add to the lifetime of a new dialog
Private Const MS_ADD_EXTRA As Integer = 5000

''
' The dialog structure
'
' @param    textLine    Array of lines of a formated chat.
' @param    X           X coord for rendering.
' @param    y           Y coord for rendering.
' @param    startTime   The time (in ms) at which the dialog was created.
' @param    lifeTime    Time (in ms) this dialog should last.
' @param    charIndex   The charIndex that created this dialog.
' @param    color       The color to be used when rendering the text.
' @param    renderable  Set to True if the chat should be rendered this frame, False otherwise
'                           (used to skip dialogs from people outside render area).
Private Type dialog
    textLine()  As String
    X           As Integer
    Y           As Integer
    Z           As Single
    startTime   As Long
    lifeTime    As Long
    CharIndex   As Integer
    color       As Long
    renderable  As Boolean
    Sube        As Byte
End Type

''
' Array if dialogs, sorted by the charIndex.
Private dialogs(MAX_DIALOGS - 1) As dialog

''
' The number of dialogs being used at the moment.
Private dialogCount As Byte

''
' the size of the font being used to render
Private userFont As tFuente

Private LastNPCCharIndex As Integer

''
' A millisecond precision timer
Private Declare Function GetTickCount Lib "kernel32" () As Long

''
' Initializes the instance of the class.

Private Sub Class_Initialize()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 07/28/07
'Initializes the dialog instance
'***************************************************
On Error GoTo ErrHandler
  
    dialogCount = 0
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Class_Initialize de clsDialogs.cls")
End Sub

''
' Clean up and release resources

Private Sub Class_Terminate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 07/28/07
'Clean up
'***************************************************
End Sub

Private Function BinarySearch(ByVal CharIndex As Integer) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 07/28/07
'Returns the index of the dialog in the list, or the negation
'of the position were it should be if not found (for binary insertion)
'**************************************************************
On Error GoTo ErrHandler
  
    Dim Min As Long
    Dim Max As Long
    Dim mid As Long
    
    Min = 0
    Max = dialogCount - 1
    
    Do While Min <= Max
        mid = (Min + Max) \ 2
        
        If dialogs(mid).CharIndex < CharIndex Then
            Min = mid + 1
        ElseIf dialogs(mid).CharIndex > CharIndex Then
            Max = mid - 1
        Else
            'We found it
            BinarySearch = mid
            Exit Function
        End If
    Loop
    
    'Not found, return the negation of the position where it should be
    '(all higher values are to the right of the list and lower values are to the left)
    BinarySearch = Not Min
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function BinarySearch de clsDialogs.cls")
End Function

''
' Creates a new dialog. If the user already had a dialog, then it's replaced,
' otherwise a new Slot is used.
'
' @param    chat        The text of the chat being spoken.
' @param    charIndex   The character saying the chat.
' @param    color       The color being used in the chat when rendered.
'
' @remarks  If there are no more free slots, and the char didn't have a previous dialog,
'           the chat is ignored.


Friend Sub CreateDialog(ByRef chat As String, ByVal CharIndex As Integer, ByVal color As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modification: 08/29/07 by Rapsodius
'Changed life time of the dialogs to a variable amount,
'depending on the number of chars
'**************************************************************
On Error GoTo ErrHandler
  
    Dim Slot As Integer
    Dim I As Long

    'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
    If charlist(CharIndex).active = False Or chat = vbNullString Then
        Call Dialogos.RemoveDialog(CharIndex)
        Exit Sub
    End If
    
    ' Display only the last NPC Message
    If charlist(CharIndex).NpcNumber > 0 Then
        If LastNPCCharIndex > 0 Then
            If charlist(LastNPCCharIndex).active Then Call Dialogos.RemoveDialog(LastNPCCharIndex)
        End If
        LastNPCCharIndex = CharIndex
    End If
    
    Slot = BinarySearch(CharIndex)
    
    If Slot < 0 Then
        If dialogCount = MAX_DIALOGS Then Exit Sub  'Out of space! Should never happen....
        
        'We need to add it. Get insertion index and move list backwards.
        Slot = Not Slot
        
        For I = dialogCount To Slot + 1 Step -1
            dialogs(I) = dialogs(I - 1)
        Next I
        
        dialogCount = dialogCount + 1
    End If
    
    With dialogs(Slot)
        .CharIndex = CharIndex
        .color = color
        .renderable = False
        .startTime = GetTickCount()
        .textLine = FormatChat(chat)
        .lifeTime = MS_ADD_EXTRA + (MS_PER_CHAR * Len(chat))
        .Sube = 8
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CreateDialog de clsDialogs.cls")
End Sub

''
' Splits a text into several lines to make it comply with the MAX_LENGTH unless it's impossible (a single word longer than MAX_LENGTH).
'
' @param    chat    The text to be formated.
'
' @return   The array of lines into which the text is splitted.
'
' @see      MAX_LENGTH

Private Function FormatChat(ByRef chat As String) As String()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 07/28/07
'Formats a dialog into different text lines.
'**************************************************************
On Error GoTo ErrHandler
  
    Dim Word As String
    Dim curPos As Long
    Dim length As Long
    Dim acumLength As Long
    Dim lineLength As Long
    Dim wordLength As Long
    Dim curLine As Long
    Dim chatLines() As String
    
    'Initialize variables
    curLine = 0
    curPos = 1
    length = Len(chat)
    acumLength = 0
    lineLength = -1
    ReDim chatLines(FieldCount(chat, 32)) As String
    
    'Start formating
    Do While acumLength < length
        Word = ReadField(curPos, chat, 32)
        
        wordLength = Len(Word)
        
        ' Is the first word of the first line? (it's the only that can start at -1)
        If lineLength = -1 Then
            chatLines(curLine) = Word
            
            lineLength = wordLength
            acumLength = wordLength
        Else
            ' Is the word too long to fit in this line?
            If lineLength + wordLength + 1 > MAX_LENGTH Then
                'Put it in the next line
                curLine = curLine + 1
                chatLines(curLine) = Word
                
                lineLength = wordLength
            Else
                'Add it to this line
                chatLines(curLine) = chatLines(curLine) & " " & Word
                
                lineLength = lineLength + wordLength + 1
            End If
            
            acumLength = acumLength + wordLength + 1
        End If
        
        'Increase to search for next word
        curPos = curPos + 1
    Loop

    
    'Resize array to fit
    ReDim Preserve chatLines(curLine) As String
    
    FormatChat = chatLines
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function FormatChat de clsDialogs.cls")
End Function

''
' Updates the position where a given dialog will be rendered.
'
' @param    X           The X coord where to start rendering.
' @param    y           The Y coord where to start rendering.
' @param    charIndex   The character whose dialog's position is being updated.

Public Sub UpdateDialogPos(ByVal X As Integer, ByVal Y As Integer, ByVal Z As Single, ByVal CharIndex As Integer)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 07/28/07
'Updates a dialog's position.
'**************************************************************
On Error GoTo ErrHandler
  
    Dim Slot As Integer
    
    Slot = BinarySearch(CharIndex)
    
    If Slot >= 0 Then
        With dialogs(Slot)
            .X = X
            .Y = Y - (UBound(.textLine()) * userFont.Tamanio)
            .Z = Z
            .renderable = True
        End With
    End If
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub UpdateDialogPos de clsDialogs.cls")
End Sub

''
' Displays all renderable dialogs on screen.

Public Sub Render()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modification: 08/29/07 by Rapsodius
'Changed the way it gets rid of older dialogs.
'**************************************************************
On Error GoTo ErrHandler
  
    Dim I As Long
    Dim J As Long
    Dim offset As Integer
    Dim removedDialogs As Long
    
    For I = 0 To dialogCount - 1

        'Decrease index to prevent jumping over a dialog
        'Crappy VB will cache the limit of the For loop, so even if it changed, it won't matter
        With dialogs(I - removedDialogs)
            If (GetTickCount() - .startTime) >= .lifeTime Then
                Call RemoveDialog(.CharIndex)
                
                removedDialogs = removedDialogs + 1
            ElseIf .renderable Then
                offset = .Y
        
                If .lifeTime >= 292 Then
                    If .Sube > 0 Then .Sube = .Sube - 1
                    If .Sube > 0 Then .Y = .Y + .Sube
                End If
                
                
                'Render it
                For J = 0 To UBound(.textLine())
                    Call DrawText(.X, offset, GetDepth(CHAT_DEPTH, .X, offset), .textLine(J), .Color, eRendererAlignmentCenterTop, FuentesJuego.FuenteBase, True)
                    offset = offset + userFont.Tamanio
                Next J
                
                'Won't render it unless it's updated in next frmae
                .renderable = False
            End If
        End With
    Next I
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub Render de clsDialogs.cls")
End Sub

''
' Removes the dialog from the requested char.
'
' @param    charIndex   The character whose dialog will be removed.

Public Sub RemoveDialog(ByVal CharIndex As Integer)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 07/28/07
'Removes the dialog from the requested char.
'**************************************************************
On Error GoTo ErrHandler
  
    Dim Slot As Integer
    Dim I As Long
    
    Slot = BinarySearch(CharIndex)
    
    If Slot < 0 Then Exit Sub
    
    For I = Slot To MAX_DIALOGS - 2
        dialogs(I) = dialogs(I + 1)
    Next I
    
    dialogCount = dialogCount - 1
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RemoveDialog de clsDialogs.cls")
End Sub

''
' Removes all dialogs from all characters.

Public Sub RemoveAllDialogs()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 07/28/07
'Removes all dialogs from all chars.
'**************************************************************
On Error GoTo ErrHandler
  
    dialogCount = 0
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RemoveAllDialogs de clsDialogs.cls")
End Sub

''
' Retrieves the font being used to render dialogs.
'
' @return   The font being used to render dialogs.

Friend Property Get Font() As tFuente
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 07/28/07
'Retrieves the font being used to render dialogs.
'**************************************************************
    Font = userFont
End Property

''
' Sets the font being used to render dialogs.
'
' @param    nValue  The new font to be used to render dialogs.

Friend Property Let Font(ByRef nValue As tFuente)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 07/28/07
'Sets the font being used to render dialogs.
'**************************************************************
    userFont = nValue
End Property
