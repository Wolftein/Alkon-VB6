VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmResources 
   Caption         =   "Seleccione directorios de recursos"
   ClientHeight    =   3420
   ClientLeft      =   2595
   ClientTop       =   1380
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cambiarBackup 
      BackColor       =   &H00FFC0C0&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cambiarMaps 
      BackColor       =   &H00FFC0C0&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton btnAceptar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cambiarDats 
      BackColor       =   &H00FFC0C0&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   375
   End
   Begin MSComDlg.CommonDialog comdiag 
      Left            =   10560
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Seleccione directorio"
      Filter          =   "Directories|*.~#~"
      InitDir         =   "App.Path"
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Por favor, seleccione la ubicacion de los siguientes recursos ya que son vitales para el funcionamiento del servidor."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1680
      TabIndex        =   10
      Top             =   240
      Width           =   8055
   End
   Begin VB.Label lblPathBackup 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   2325
      Width           =   8175
   End
   Begin VB.Label lblPathMaps 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   1710
      Width           =   8175
   End
   Begin VB.Label Label4 
      Caption         =   "WorldBackup:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "MAPS:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblPathDats 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   1110
      Width           =   8175
   End
   Begin VB.Label Label1 
      Caption         =   "DATS:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "frmResources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ServerIniManager As clsIniManager

Private Sub btnAceptar_Click()
    Dim ValidationMessage As String
    
    If Not AllPathsConfigured(ValidationMessage) Then
        Call MsgBox(ValidationMessage)
        Exit Sub
    End If

    Call WriteVar(IniPath & "Server.ini", "RECURSOS", "Dats", lblPathDats.Caption)
    Call WriteVar(IniPath & "Server.ini", "RECURSOS", "Maps", lblPathMaps.Caption)
    Call WriteVar(IniPath & "Server.ini", "RECURSOS", "WorldBackup", lblPathBackup.Caption)

    Unload frmResources
    Call Main
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim PathsOk As Boolean
    Dim ValidationMessage As String
        
    If Not AllPathsConfigured(ValidationMessage) Then
        If MsgBox(ValidationMessage & vbNewLine & "Si presiona Ok el servidor se cerrara automaticamente.", vbOKCancel) = vbOK Then
            End
        Else
            Cancel = True
        End If
    End If
    
End Sub
Private Sub OpenCommonDialog(ByVal Tipo As String)
    comdiag.DialogTitle = "Seleccione un directorio"
    comdiag.InitDir = App.Path
    comdiag.FileName = "Seleccione un directorio para " & Tipo
    comdiag.flags = cdlOFNNoValidate + cdlOFNHideReadOnly
    comdiag.Filter = "Directories|*.~#~"
    comdiag.ShowSave
End Sub

Private Sub cambiarBackup_Click()
    Call OpenCommonDialog("el world backup")
    If Err <> 32755 Then
        lblPathBackup.Caption = CurDir & "\"
    End If
End Sub

Private Sub cambiarDats_Click()
    Call OpenCommonDialog("los dats")
    If Err <> 32755 Then
        lblPathDats.Caption = CurDir & "\"
    End If
End Sub

Private Sub cambiarMaps_Click()
    Call OpenCommonDialog("los maps")
    If Err <> 32755 Then
        lblPathMaps.Caption = CurDir & "\"
    End If
End Sub

Private Sub Form_Load()
    With ServerConfiguration.ResourcesPaths
        lblPathBackup.Caption = .WorldBackup
        lblPathDats.Caption = .Dats
        lblPathMaps.Caption = .Maps
    End With
    Set ServerIniManager = New clsIniManager
    ServerIniManager.Initialize (IniPath & "Server.ini")
End Sub

Private Function AllPathsConfigured(ByRef ValidationMessage As String) As Boolean
    AllPathsConfigured = False
    
    If Not General.FileExist(lblPathDats.Caption, vbDirectory) Or lblPathDats.Caption = "" Then
        ValidationMessage = "El path de los dats no es un directorio valido"
        Exit Function
    End If
    
    If Not General.FileExist(lblPathMaps.Caption, vbDirectory) Or lblPathMaps.Caption = "" Then
        ValidationMessage = "El path de los maps no es un directorio valido"
        Exit Function
    End If
    
    If Not General.FileExist(lblPathBackup.Caption, vbDirectory) Or lblPathBackup.Caption = "" Then
        ValidationMessage = "El path del world backup no es un directorio valido"
        Exit Function
    End If
    
    AllPathsConfigured = True
End Function

