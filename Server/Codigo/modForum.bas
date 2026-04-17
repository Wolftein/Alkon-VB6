Attribute VB_Name = "modForum"
'Argentum Online 0.14.0
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez


Option Explicit

Public Const FORO_REAL_ID As String = "REAL"
Public Const FORO_CAOS_ID As String = "CAOS"

Public Type tPost
    sTitulo As String
    sPost As String
    Autor As String
End Type

Public Type tForo
    vsPost() As tPost
    vsAnuncio() As tPost
    CantPosts As Byte
    CantAnuncios As Byte
    ID As String
End Type

Private NumForos As Integer
Private Foros() As tForo


Public Sub AddForum(ByVal sForoID As String)
'***************************************************
'Author: ZaMa
'Last Modification: 22/02/2010
'Adds a forum to the list and fills it.
'***************************************************
On Error GoTo ErrHandler
  
    Dim ForumPath As String
    Dim PostPath As String
    Dim PostIndex As Integer
    Dim FileIndex As Integer
    
    NumForos = NumForos + 1
    ReDim Preserve Foros(1 To NumForos) As tForo
    
    ForumPath = App.Path & "\foros\" & sForoID & ".for"
    
    With Foros(NumForos)
        
        ReDim .vsPost(1 To Constantes.MaxMensajesForo) As tPost
        ReDim .vsAnuncio(1 To Constantes.MaxAnunciosForo) As tPost
        
        .ID = sForoID
        
        If FileExist(ForumPath, vbNormal) Then
            .CantPosts = Val(GetVar(ForumPath, "INFO", "CantMSG"))
            .CantAnuncios = Val(GetVar(ForumPath, "INFO", "CantAnuncios"))
            
            ' Cargo posts
            For PostIndex = 1 To .CantPosts
                FileIndex = FreeFile
                PostPath = App.Path & "\foros\" & sForoID & PostIndex & ".for"

                Open PostPath For Input Shared As #FileIndex
                
                ' Titulo
                Input #FileIndex, .vsPost(PostIndex).sTitulo
                ' Autor
                Input #FileIndex, .vsPost(PostIndex).Autor
                ' Mensaje
                Input #FileIndex, .vsPost(PostIndex).sPost
                
                Close #FileIndex
            Next PostIndex
            
            ' Cargo anuncios
            For PostIndex = 1 To .CantAnuncios
                FileIndex = FreeFile
                PostPath = App.Path & "\foros\" & sForoID & PostIndex & "a.for"

                Open PostPath For Input Shared As #FileIndex
                
                ' Titulo
                Input #FileIndex, .vsAnuncio(PostIndex).sTitulo
                ' Autor
                Input #FileIndex, .vsAnuncio(PostIndex).Autor
                ' Mensaje
                Input #FileIndex, .vsAnuncio(PostIndex).sPost
                
                Close #FileIndex
            Next PostIndex
        End If
        
    End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AddForum de modForum.bas")
End Sub

Public Function GetForumIndex(ByRef sForoID As String) As Integer
'***************************************************
'Author: ZaMa
'Last Modification: 22/02/2010
'Returns the forum index.
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim ForumIndex As Integer
    
    For ForumIndex = 1 To NumForos
        If Foros(ForumIndex).ID = sForoID Then
            GetForumIndex = ForumIndex
            Exit Function
        End If
    Next ForumIndex
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function GetForumIndex de modForum.bas")
End Function

Public Sub AddPost(ByVal ForumIndex As Integer, ByRef Post As String, ByRef Autor As String, _
                   ByRef Titulo As String, ByVal bAnuncio As Boolean)
'***************************************************
'Author: ZaMa
'Last Modification: 22/02/2010
'Saves a new post into the forum.
'***************************************************
On Error GoTo ErrHandler
  

    With Foros(ForumIndex)
        
        If bAnuncio Then
            If .CantAnuncios < Constantes.MaxAnunciosForo Then _
                .CantAnuncios = .CantAnuncios + 1
            
            Call MoveArray(ForumIndex, bAnuncio)
            
            ' Agrego el anuncio
            With .vsAnuncio(1)
                .sTitulo = Titulo
                .Autor = Autor
                .sPost = Post
            End With
            
        Else
            If .CantPosts < Constantes.MaxMensajesForo Then _
                .CantPosts = .CantPosts + 1
                
            Call MoveArray(ForumIndex, bAnuncio)
            
            ' Agrego el post
            With .vsPost(1)
                .sTitulo = Titulo
                .Autor = Autor
                .sPost = Post
            End With
        
        End If
    End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub AddPost de modForum.bas")
End Sub

Public Sub SaveForums()
'***************************************************
'Author: ZaMa
'Last Modification: 22/02/2010
'Saves all forums into disk.
'***************************************************
On Error GoTo ErrHandler
  
    Dim ForumIndex As Integer

    For ForumIndex = 1 To NumForos
        Call SaveForum(ForumIndex)
    Next ForumIndex
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SaveForums de modForum.bas")
End Sub


Private Sub SaveForum(ByVal ForumIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 22/02/2010
'Saves a forum into disk.
'***************************************************
On Error GoTo ErrHandler
  

    Dim PostIndex As Integer
    Dim FileIndex As Integer
    Dim PostPath As String
    
    Call CleanForum(ForumIndex)
    
    With Foros(ForumIndex)
        
        ' Guardo info del foro
        Call WriteVar(App.Path & "\Foros\" & .ID & ".for", "INFO", "CantMSG", .CantPosts)
        Call WriteVar(App.Path & "\Foros\" & .ID & ".for", "INFO", "CantAnuncios", .CantAnuncios)
        
        ' Guardo posts
        For PostIndex = 1 To .CantPosts
            
            PostPath = App.Path & "\Foros\" & .ID & PostIndex & ".for"
            FileIndex = FreeFile()
            Open PostPath For Output As FileIndex
            
            With .vsPost(PostIndex)
                Print #FileIndex, .sTitulo
                Print #FileIndex, .Autor
                Print #FileIndex, .sPost
            End With
            
            Close #FileIndex
            
        Next PostIndex
        
        ' Guardo Anuncios
        For PostIndex = 1 To .CantAnuncios
            
            PostPath = App.Path & "\Foros\" & .ID & PostIndex & "a.for"
            FileIndex = FreeFile()
            Open PostPath For Output As FileIndex
            
            With .vsAnuncio(PostIndex)
                Print #FileIndex, .sTitulo
                Print #FileIndex, .Autor
                Print #FileIndex, .sPost
            End With
            
            Close #FileIndex

        Next PostIndex
        
    End With
    
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SaveForum de modForum.bas")
End Sub

Public Sub CleanForum(ByVal ForumIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 22/02/2010
'Cleans a forum from disk.
'***************************************************
On Error GoTo ErrHandler
  
    Dim PostIndex As Integer
    Dim NumPost As Integer
    Dim ForumPath As String

    With Foros(ForumIndex)
    
        ' Elimino todo
        ForumPath = App.Path & "\Foros\" & .ID & ".for"
        If FileExist(ForumPath, vbNormal) Then
    
            NumPost = Val(GetVar(ForumPath, "INFO", "CantMSG"))
            
            ' Elimino los post viejos
            For PostIndex = 1 To NumPost
                Kill App.Path & "\Foros\" & .ID & PostIndex & ".for"
            Next PostIndex
            
            
            NumPost = Val(GetVar(ForumPath, "INFO", "CantAnuncios"))
            
            ' Elimino los post viejos
            For PostIndex = 1 To NumPost
                Kill App.Path & "\Foros\" & .ID & PostIndex & "a.for"
            Next PostIndex
            
            
            ' Elimino el foro
            Kill App.Path & "\Foros\" & .ID & ".for"
    
        End If
    End With

  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CleanForum de modForum.bas")
End Sub

Public Function SendPosts(ByVal UserIndex As Integer, ByRef ForoID As String) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 22/02/2010
'Sends all the posts of a required forum
'***************************************************
On Error GoTo ErrHandler
  
    
    Dim ForumIndex As Integer
    Dim PostIndex As Integer
    Dim bEsGm As Boolean
    
    ForumIndex = GetForumIndex(ForoID)

    If ForumIndex > 0 Then

        With Foros(ForumIndex)
            
            ' Send General posts
            For PostIndex = 1 To .CantPosts
                With .vsPost(PostIndex)
                    Call WriteAddForumMsg(UserIndex, eForumMsgType.ieGeneral, .sTitulo, .Autor, .sPost)
                End With
            Next PostIndex
            
            ' Send Sticky posts
            For PostIndex = 1 To .CantAnuncios
                With .vsAnuncio(PostIndex)
                    Call WriteAddForumMsg(UserIndex, eForumMsgType.ieGENERAL_STICKY, .sTitulo, .Autor, .sPost)
                End With
            Next PostIndex
            
        End With
        
        bEsGm = EsGm(UserIndex)
        
        ' Caos?
        If esCaos(UserIndex) Or bEsGm Then
            
            ForumIndex = GetForumIndex(FORO_CAOS_ID)
            
            With Foros(ForumIndex)
                
                ' Send General Caos posts
                For PostIndex = 1 To .CantPosts
                
                    With .vsPost(PostIndex)
                        Call WriteAddForumMsg(UserIndex, eForumMsgType.ieCAOS, .sTitulo, .Autor, .sPost)
                    End With
                    
                Next PostIndex
                
                ' Send Sticky posts
                For PostIndex = 1 To .CantAnuncios
                    With .vsAnuncio(PostIndex)
                        Call WriteAddForumMsg(UserIndex, eForumMsgType.ieCAOS_STICKY, .sTitulo, .Autor, .sPost)
                    End With
                Next PostIndex
                
            End With
        End If
            
        ' Caos?
        If esArmada(UserIndex) Or bEsGm Then
            
            ForumIndex = GetForumIndex(FORO_REAL_ID)
            
            With Foros(ForumIndex)
                
                ' Send General Real posts
                For PostIndex = 1 To .CantPosts
                
                    With .vsPost(PostIndex)
                        Call WriteAddForumMsg(UserIndex, eForumMsgType.ieREAL, .sTitulo, .Autor, .sPost)
                    End With
                    
                Next PostIndex
                
                ' Send Sticky posts
                For PostIndex = 1 To .CantAnuncios
                    With .vsAnuncio(PostIndex)
                        Call WriteAddForumMsg(UserIndex, eForumMsgType.ieREAL_STICKY, .sTitulo, .Autor, .sPost)
                    End With
                Next PostIndex
                
            End With
        End If
        
        SendPosts = True
    End If
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function SendPosts de modForum.bas")
End Function

Public Function EsAnuncio(ByVal ForumType As Byte) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 22/02/2010
'Returns true if the post is sticky.
'***************************************************
On Error GoTo ErrHandler
  
    Select Case ForumType
        Case eForumMsgType.ieCAOS_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieGENERAL_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieREAL_STICKY
            EsAnuncio = True
            
    End Select
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function EsAnuncio de modForum.bas")
End Function

Public Function ForumAlignment(ByVal yForumType As Byte) As Byte
'***************************************************
'Author: ZaMa
'Last Modification: 01/03/2010
'Returns the forum alignment.
'***************************************************
On Error GoTo ErrHandler
  
    Select Case yForumType
        Case eForumMsgType.ieCAOS, eForumMsgType.ieCAOS_STICKY
            ForumAlignment = eForumType.ieCAOS
            
        Case eForumMsgType.ieGeneral, eForumMsgType.ieGENERAL_STICKY
            ForumAlignment = eForumType.ieGeneral
            
        Case eForumMsgType.ieREAL, eForumMsgType.ieREAL_STICKY
            ForumAlignment = eForumType.ieREAL
            
    End Select
    
  
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ForumAlignment de modForum.bas")
End Function

Public Sub ResetForums()
'***************************************************
'Author: ZaMa
'Last Modification: 22/02/2010
'Resets forum info
'***************************************************
On Error GoTo ErrHandler
  
    ReDim Foros(1 To 1) As tForo
    NumForos = 0
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub ResetForums de modForum.bas")
End Sub

Private Sub MoveArray(ByVal ForumIndex As Integer, ByVal Sticky As Boolean)
On Error GoTo ErrHandler
  
Dim I As Long

With Foros(ForumIndex)
    If Sticky Then
        For I = .CantAnuncios To 2 Step -1
            .vsAnuncio(I).sTitulo = .vsAnuncio(I - 1).sTitulo
            .vsAnuncio(I).sPost = .vsAnuncio(I - 1).sPost
            .vsAnuncio(I).Autor = .vsAnuncio(I - 1).Autor
        Next I
    Else
        For I = .CantPosts To 2 Step -1
            .vsPost(I).sTitulo = .vsPost(I - 1).sTitulo
            .vsPost(I).sPost = .vsPost(I - 1).sPost
            .vsPost(I).Autor = .vsPost(I - 1).Autor
        Next I
    End If
End With
  
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub MoveArray de modForum.bas")
End Sub
