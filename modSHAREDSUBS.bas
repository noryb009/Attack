Attribute VB_Name = "modSHAREDSUBS"
Public Enum monsterNames
    greenMonster
    blackMonster
    bat
    tree
    cloud
    rabbit
    ladybug
    knightSword
    knightFlail
    knightHorse
    dragon
End Enum

Global Const landHEIGHT = 376
Global Const numberOfMonsters = 11
Global cmontypeMONSTERINFO(0 To numberOfMonsters - 1) As New clsMONSTERTYPE
Global Const flailSIZEPX = 14

Global Const windowX = 700
Global Const windowY = 500

Function safeADDLONG(lNUMBER1 As Long, lNUMBER2 As Long) As Long
    Dim dblOUTPUT As Double
    dblOUTPUT = lNUMBER1
    dblOUTPUT = dblOUTPUT + lNUMBER2
    If dblOUTPUT < 2147483647 Then
        safeADDLONG = dblOUTPUT
    Else
        safeADDLONG = 2147483647
    End If
End Function

Sub loadMONSTERINFO()
    loadONEMONSTERINFO greenMonster, "monster0", 9, 25, 1, 2, -1, 1, 0, 2
    loadONEMONSTERINFO blackMonster, "monster1", 9, 25, 2, 5, -1, 1, 1, 2
    loadONEMONSTERINFO bat, "monster2", 10, 11, 1, 3, 150, 1.5, 0, 2
    loadONEMONSTERINFO tree, "monster3", 26, 50, 10, 8, -1, 0.4, 1, 5
    loadONEMONSTERINFO cloud, "monster4", 43, 28, 3, 5, 10, 1, 1, 3
    loadONEMONSTERINFO rabbit, "monster5", 17, 34, 4, 3, -1, 2, 1, 3
    loadONEMONSTERINFO ladybug, "monster6", 13, 7, 4, 2, -1, 2.5, 1, 2
    loadONEMONSTERINFO knightSword, "knight", 21, 51, 10, 20, -1, 0.5, 1, 4
    loadONEMONSTERINFO knightFlail, "knightFlail", 33, 51, 15, 35, -1, 0.5, 1, 6
    loadONEMONSTERINFO knightHorse, "knightHorse", 92, 43, 8, 20, -1, 3, 1, 8
    loadONEMONSTERINFO dragon, "dragon", 91, 53, 50, 200, 200, 0.3, 0, 10
End Sub
