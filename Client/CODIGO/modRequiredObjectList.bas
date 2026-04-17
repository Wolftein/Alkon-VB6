Attribute VB_Name = "modRequiredObjectList"
'@Folder("Quest")
Option Explicit

Public Type RequiredObjectListItem
    ObjIndex As Integer
    Quantity As Long
    RequiredQuantity As Long
End Type

Public Type RequiredObjectList
    Items() As RequiredObjectListItem
    ItemsCount As Integer
    IsComplete As Boolean
End Type
Public Function RequiredObjectListCreateCompleted() As RequiredObjectList
    Dim List As RequiredObjectList
    Dim ItemIndex As Long
    
    With List
        .ItemsCount = 0
        .IsComplete = True
    End With
    
    RequiredObjectListCreateCompleted = List
End Function
Public Function RequiredObjectListCreate(ByRef RequiredItems() As RequiredObjectListItem, ByVal ItemsCount As Integer) As RequiredObjectList
    Dim List As RequiredObjectList
    Dim ItemIndex As Long
    With List
        .Items = RequiredItems
        .ItemsCount = ItemsCount
         For ItemIndex = 0 To .ItemsCount - 1
            .Items(ItemIndex).Quantity = 0
        Next ItemIndex
        
        .IsComplete = False
    End With
    
    RequiredObjectListCreate = List
End Function
Public Function RequiredObjectListTryAdd(ByRef List As RequiredObjectList, ByVal ObjIndex As Integer, ByVal Quantity As Long, ByRef RestQuantity As Long) As Boolean
    Dim ItemIndex As Long

    With List
        For ItemIndex = 0 To .ItemsCount - 1
             If ObjIndex = .Items(ItemIndex).ObjIndex Then
                If .Items(ItemIndex).Quantity = .Items(ItemIndex).RequiredQuantity Then
                    RequiredObjectListTryAdd = False
                    Exit Function
                End If
                .Items(ItemIndex).Quantity = .Items(ItemIndex).Quantity + Quantity
                
                If .Items(ItemIndex).Quantity > .Items(ItemIndex).RequiredQuantity Then
                    RestQuantity = .Items(ItemIndex).Quantity - .Items(ItemIndex).RequiredQuantity
                    .Items(ItemIndex).Quantity = .Items(ItemIndex).RequiredQuantity
                Else
                    RestQuantity = 0
                End If
                
                RequiredObjectListTryAdd = True
                Exit Function
            End If
        Next ItemIndex
        
        RequiredObjectListTryAdd = False
    End With
End Function

Public Function RequiredObjectListIsComplete(ByRef List As RequiredObjectList) As Boolean
    Dim ItemIndex As Long

    With List
        If List.IsComplete Then
            RequiredObjectListIsComplete = True
            Exit Function
        End If
        
        For ItemIndex = 0 To .ItemsCount - 1
            If .Items(ItemIndex).Quantity < .Items(ItemIndex).RequiredQuantity Then
                RequiredObjectListIsComplete = False
                Exit Function
            End If
        Next ItemIndex
        
        List.IsComplete = True
        RequiredObjectListIsComplete = True
    End With

End Function
