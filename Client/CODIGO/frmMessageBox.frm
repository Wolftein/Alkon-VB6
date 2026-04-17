VERSION 5.00
Begin VB.Form frmMessageBox 
   Appearance      =   0  'Flat
   BackColor       =   &H00292929&
   BorderStyle     =   0  'None
   Caption         =   "Message"
   ClientHeight    =   2625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3975
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   175
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label LblMessage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
   Begin VB.Image ImgAcceptButton 
      Height          =   495
      Left            =   1200
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Image ImgFrameCornerBottomRight 
      Height          =   255
      Left            =   3600
      Top             =   1440
      Width           =   255
   End
   Begin VB.Image ImgFrameCornerBottomLeft 
      Height          =   255
      Left            =   120
      Top             =   1440
      Width           =   255
   End
   Begin VB.Image ImgFrameCornerTopRight 
      Height          =   255
      Left            =   3600
      Top             =   120
      Width           =   255
   End
   Begin VB.Image ImgFrameCornerTopLeft 
      Height          =   255
      Left            =   120
      Top             =   120
      Width           =   255
   End
   Begin VB.Image ImgFrameBottom 
      Height          =   135
      Left            =   720
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Image ImgFrameTop 
      Height          =   135
      Left            =   720
      Top             =   120
      Width           =   2535
   End
   Begin VB.Image ImgFrameRight 
      Height          =   735
      Left            =   3720
      Top             =   480
      Width           =   135
   End
   Begin VB.Image ImgFrameLeft 
      Height          =   735
      Left            =   120
      Top             =   480
      Width           =   135
   End
End
Attribute VB_Name = "frmMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, _
                                                                lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_CALCRECT = &H400&
Private Const DT_WORDBREAK = &H10&

Public LastButtonPressed As clsGraphicalButton

Private clsFormulario As clsFormMovementManager
Private cButtonAccept As clsGraphicalButton

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
 

Private StringMessage As String
Const MinWidth As Long = 6600
Const MinHeight As Long = 2700

Private Sub RepositionAndResizeControls()
    Dim OldScaleMode As Integer
    OldScaleMode = Me.ScaleMode
    
    Me.ScaleMode = vbTwips
    
    LblMessage.Width = Me.Width - (LblMessage.Left * 2) - 10

    ImgAcceptButton.Top = Me.ScaleHeight - ImgAcceptButton.Height - 200
    
    ImgAcceptButton.Left = (Me.ScaleWidth / 2) - (ImgAcceptButton.Width / 2)
  
    Me.ScaleMode = OldScaleMode
End Sub

Public Sub ShowMessage(ByRef Message As String, Optional ByRef ParentForm As Form = Nothing)

    StringMessage = Message
    
    If Not ParentForm Is Nothing Then
        Me.Show , ParentForm
    Else
        Me.Show , Screen.ActiveForm
    End If
    
    
    Call RepositionAndResizeControls
    
    Call SetLabelCaption(LblMessage, Message, Me)
    
    Call SetFormSize(Len(Message))
    
End Sub


Private Sub Form_Load()
    ' This is required as the text will be drawn in the form and not in the label.
    ' And both fonts need to be the same.
    Me.Font = LblMessage.Font
    Me.FontBold = LblMessage.FontBold
    Me.FontItalic = LblMessage.FontBold
    Me.FontName = LblMessage.FontName
    Me.FontSize = LblMessage.FontSize
    Me.FontStrikethru = LblMessage.FontStrikethru
    Me.FontUnderline = LblMessage.FontUnderline
    
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me, , False
    
    Call InitializeImageControls
    
    Call modCustomCursors.SetFormCursorDefault(Me)
End Sub

Private Sub InitializeImageControls()
    Dim GrhPath As String
    GrhPath = DirInterfaces & SELECTED_UI
    
    
    
    ImgFrameTop.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "FrameBarraArriba.JPG")
    ImgFrameBottom.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "FrameBarraAbajo.JPG")
    ImgFrameLeft.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "FrameBarraIzquierda.JPG")
    ImgFrameRight.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "FrameBarraDerecha.JPG")
    
    ImgFrameCornerTopLeft.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "FrameArribaIzquierda.JPG")
    ImgFrameCornerTopRight.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "FrameArribaDerecha.JPG")
    ImgFrameCornerBottomLeft.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "FrameAbajoIzquierda.JPG")
    ImgFrameCornerBottomRight.Picture = LoadPicture(DirInterfaces & SELECTED_UI & "FrameAbajoDerecha.JPG")
    
    Set LastButtonPressed = New clsGraphicalButton
    Set cButtonAccept = New clsGraphicalButton
    
    
    Call cButtonAccept.Initialize(ImgAcceptButton, GrhPath & "BotonAceptar.jpg", _
                                    GrhPath & "BotonAceptar.jpg", _
                                    GrhPath & "BotonAceptar.jpg", Me)

End Sub

Private Sub Form_Resize()
 
    Call AdjustFrame
    
    Call RepositionAndResizeControls
    
    Call SetLabelCaption(LblMessage, StringMessage, Me)
    
End Sub

Private Sub AdjustFrame()
    ImgFrameTop.Left = 0
    ImgFrameTop.Top = 0
    
    ImgFrameLeft.Left = 0
    ImgFrameLeft.Top = 0
    
    ImgFrameCornerTopLeft.Left = 0
    ImgFrameCornerTopLeft.Top = 0
    
    ImgFrameCornerBottomLeft.Left = 0
    ImgFrameCornerBottomLeft.Top = Me.ScaleHeight - ImgFrameCornerBottomLeft.Height
    
    ImgFrameCornerTopRight.Left = Me.ScaleWidth - ImgFrameCornerTopRight.Width
    ImgFrameCornerTopRight.Top = 0
    
    ImgFrameRight.Left = Me.ScaleWidth - ImgFrameRight.Width
    ImgFrameRight.Top = 0
        
    ImgFrameCornerBottomRight.Left = Me.ScaleWidth - ImgFrameCornerBottomRight.Width
    ImgFrameCornerBottomRight.Top = Me.ScaleHeight - ImgFrameCornerBottomRight.Height

    ImgFrameBottom.Left = 0
    ImgFrameBottom.Top = Me.ScaleHeight - ImgFrameBottom.Height
    
        
End Sub


Private Sub SetFormSize(ByVal MessageLength As Long)
    
    Dim OldScaleMode As Integer
    OldScaleMode = Me.ScaleMode
    '
    frmMessageBox.ScaleMode = vbTwips
   
    Dim CantLines As Long
    CantLines = Round((MessageLength / 60) + 0.5)
   
    frmMessageBox.Height = 1000 + LblMessage.Height + 300 + ImgAcceptButton.Height + 50
   
   
    frmMessageBox.ScaleMode = OldScaleMode

End Sub

Public Sub SetLabelCaption(Lbl As Label, ByVal Caption As String, ByRef Form As Form)
    Dim Rct As RECT
    Dim OldScaleMode As Long
    Dim Border As Long
    OldScaleMode = Form.ScaleMode
    
    'Change the scalemode to Pixels to simplify the calculations
    Form.ScaleMode = vbPixels
    If Lbl.BorderStyle <> vbBSNone Then
        If Lbl.Appearance = 1 Then
            '3D border
            Border = 4
        Else
            Border = 2
        End If
    End If
    Rct.Right = Lbl.Width - Border
    
    Dim TextSize As Long
    TextSize = DrawText(Form.hdc, Caption, Len(Caption), Rct, DT_WORDBREAK + DT_CALCRECT)
  
    Lbl.Height = TextSize + Border
    Lbl.Caption = Caption
    
    'Restore the ScaleMode
    Form.ScaleMode = OldScaleMode
End Sub

Private Sub ImgAcceptButton_Click()
    Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

