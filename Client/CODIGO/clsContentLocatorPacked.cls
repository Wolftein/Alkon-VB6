VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContentLocatorPacked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements Aurora_Engine.Content_Locator_

Private Const BASE_URL As String = "\Graficos\" ' TODO: Nightw - Change this to "\Resources\"

Private Function Content_Locator__Open(ByVal fileName As String) As Aurora_Engine.Memory_Chunk

    Dim Header As INFOHEADER
    
    If Get_InfoHeader(App.path & BASE_URL, fileName, Header) Then
        Dim Bytes() As Byte
        Call Extract_File(App.path & BASE_URL, Header, Bytes)

        ' TODO: Zero Memory Copy
        Set Content_Locator__Open = New Aurora_Engine.Memory_Chunk
        Call Content_Locator__Open.Allocate(UBound(Bytes))
        Call Content_Locator__Open.Copy(Bytes, UBound(Bytes), 0)
    End If
 
End Function

Private Sub Content_Locator__Delete(ByVal Path As String)
End Sub

Private Sub Content_Locator__Write(ByVal Path As String, ByVal Data As Aurora_Engine.BinaryReader_)
End Sub
