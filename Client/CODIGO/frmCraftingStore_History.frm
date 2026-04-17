VERSION 5.00
Begin VB.Form frmCraftingStore_History 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Historial de ventas"
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleMode       =   0  'User
   ScaleWidth      =   8508.63
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtLog 
      BackColor       =   &H00292929&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   7935
   End
   Begin VB.Image imgAccept 
      Height          =   330
      Left            =   3600
      Top             =   4560
      Width           =   1050
   End
   Begin VB.Label lblGold 
      Appearance      =   0  'Flat
      BackColor       =   &H00292929&
      BackStyle       =   0  'Transparent
      Caption         =   "999999999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   360
      Width           =   1245
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00292929&
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   360
      Width           =   405
   End
   Begin VB.Label lblSales 
      Appearance      =   0  'Flat
      BackColor       =   &H00292929&
      BackStyle       =   0  'Transparent
      Caption         =   "999999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   360
      Width           =   765
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00292929&
      BackStyle       =   0  'Transparent
      Caption         =   "Ventas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   645
   End
End
Attribute VB_Name = "frmCraftingStore_History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SalesQuantity As Long
Public GoldAquired As Double

Private cButtonClose As clsGraphicalButton
Public LastButtonPressed As clsGraphicalButton

Private cForm As clsFormMovementManager


Private Sub CloseWindow()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Call CloseWindow
End Sub

Public Sub AddCraftedItemLog(ByVal ItemNumber As String, ByVal ItemQuantity As Integer, ByVal ConstructionPrice As Double, ByRef BuyerName As String)
    If SalesQuantity = 0 Then
        txtLog.text = ""
    End If

    SalesQuantity = SalesQuantity + ItemQuantity
    GoldAquired = GoldAquired + ConstructionPrice
    
    lblSales.Caption = SalesQuantity
    lblGold.Caption = GoldAquired
    
    txtLog.text = txtLog.text & GameMetadata.Objs(ItemNumber).Name & " por " & ConstructionPrice & " de oro a " & BuyerName & vbCrLf
    txtLog.Locked = True
    
End Sub


Public Sub CleanControls()
    SalesQuantity = 0
    GoldAquired = 0
    
    txtLog.text = ""
    lblGold.Caption = 0
    lblSales.Caption = 0
    txtLog.Locked = True
        
End Sub

Private Sub Form_Load()
    Call CleanControls
    
    Call InitUI
    
    Call modCustomCursors.SetFormCursorDefault(Me)
End Sub

Public Sub InitUI()

    Dim ImgPath As String

    Set cForm = New clsFormMovementManager
    cForm.Initialize Me, , False
    
    Set LastButtonPressed = New clsGraphicalButton
        
    Set cButtonClose = New clsGraphicalButton
    
    ImgPath = DirInterfaces & SELECTED_UI
    Set Me.Picture = LoadPicture(ImgPath & "VentanaSelfWorker_History.jpg")
    
    Call cButtonClose.Initialize(imgAccept, ImgPath & "BotonAceptar.jpg", ImgPath & "BotonAceptar.jpg", ImgPath & "BotonAceptar.jpg", Me, ImgPath & "BotonAceptar.jpg")
        
End Sub

Private Sub imgAccept_Click()
    Me.Visible = False
End Sub
