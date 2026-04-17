Attribute VB_Name = "modDropSystem"
Option Explicit

Private Type tDropItem
    ObjIndex As Integer
    Amount As Long
End Type

Public Type tDropData
    NumItems As Integer
    Item() As tDropItem
End Type

Public DropData() As tDropData

Public Sub LoadDropData()
On Error GoTo ErrHandler

    Dim I As Long
    Dim Item As Integer
    Dim NumDrops As Long
    Dim Tmp As String
    Dim TmpArray() As String
    
    NumDrops = Val(GetVar(DatPath & "Drops.dat", "INIT", "NumDrops"))
    ReDim DropData(1 To NumDrops) As tDropData
    
    For I = 1 To NumDrops
        With DropData(I)
            .NumItems = Val(GetVar(DatPath & "Drops.dat", "Drop" & I, "NumItems"))
            If .NumItems > 0 Then
                ReDim .Item(1 To .NumItems) As tDropItem
                For Item = 1 To .NumItems
                    Tmp = GetVar(DatPath & "Drops.dat", "Drop" & I, "Item" & Item)
                    .Item(Item).ObjIndex = Val(ReadField(1, Tmp, Asc("-")))
                    .Item(Item).Amount = Val(ReadField(2, Tmp, Asc("-")))
                Next Item
            End If
        End With
    Next I
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error " & Err.Number & " (" & Err.Description & ") en Sub LoadDropData del Módulo modDropSystem")
End Sub
