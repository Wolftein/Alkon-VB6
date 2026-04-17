VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsButtonEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event GraphicalButtonClick(gb As clsGraphicalButton)
Public Event GraphicalInventoryClick(gb As clsGraphicalInventory)

Public Sub OnGraphicalButtonClick(gb As clsGraphicalButton)
    RaiseEvent GraphicalButtonClick(gb)
End Sub

Public Sub OnGraphicalInventoryClick(gb As clsGraphicalInventory)
    RaiseEvent GraphicalInventoryClick(gb)
End Sub
