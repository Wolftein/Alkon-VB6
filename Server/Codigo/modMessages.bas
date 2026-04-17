Attribute VB_Name = "modMessages"
Option Explicit

Public Enum eMessageId
    None = 0
    Spell_Hits_Npc = 1
    Char_Killed_By_User = 2
    Char_Hit = 3
    Cant_Quit_Paralized = 4

    Guild_No_Permission = 5
    
    Guild_Invitation_Limit_Reached = 6
    Guild_Invitation_User_Offline = 7
    Guild_Invitation_User_Already_Has_Guild = 8
    Guild_Invitation_User_In_Other_Faction = 9
    Guild_Invitation_User_Already_Invitated = 10
    Guild_Invitation_Sent = 11
    Guild_Invitation_User_Invitated = 12
End Enum


Public Enum eMessageParameterType
    Number = 0
    text = 1
End Enum


Public MessageManager As New clsMessageManager
