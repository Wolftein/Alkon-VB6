Attribute VB_Name = "ModPunishments"
Option Explicit

Public Enum ePunishmentSubType
    Jail = 1
    Ban = 2
    Warning = 3
End Enum

Public Type tPunishmentRule
    Count As Integer
    Severity As Integer
End Type


Public Type tPunishmentType
    Id As String
    Name As String
    BaseType As Byte
    Rules() As tPunishmentRule
End Type

Public punishmentList() As tPunishmentType

