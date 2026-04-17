Attribute VB_Name = "modTournament"
Option Explicit

' NOTA: Dejo esto aparte, sino se provocan referencias circulares..

' Tournament
Public Const MAX_ARENAS As Byte = 5

Public Enum eTournamentEdit
    ieMaxCompetitor
    ieMaxLevel
    ieMinLevel
    ieRequiredGold
    ieForbiddenItems
    iePermitedClass
    ieNumRoundsToWin
    ieKillAfterLoose
    ieWaitingMap
    ieArenaPosition
    ieFinalMap
    ieSaveConfig

    ieLastOption ' Number of edit options
End Enum

Public Type tArena
    ' Default positions
    Map As Integer
    UserPos1 As Position
    UserPos2 As Position
End Type
'
Private Type tTournament

    ' Competitors
    MaxCompetitors As Byte
    CompetitorsList() As String

    ' Restrictions
    MinLevel As Byte
    MaxLevel As Byte
    RequiredGold As Long

    NumForbiddenItems As Byte
    ForbiddenItem() As Integer

    NumPermitedClass As Byte
    PermitedClass() As Byte

    ' Aditionals
    NumRoundsToWin As Byte
    KillAfterLoose As Byte

    ' Positions
    WaitingMap As WorldPos
    FinalMap As WorldPos
    Arenas(1 To MAX_ARENAS) As tArena
End Type

' Public holder
Public Tournament As tTournament

