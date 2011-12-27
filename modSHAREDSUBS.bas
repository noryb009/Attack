Attribute VB_Name = "modSHAREDSUBS"
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

Global Const numberOfMonsters = 11 ' number of types of monsters
Public Enum monsterNames ' names of monsters
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

Global Const landHEIGHT = 376 ' height of land
Global cmontypeMONSTERINFO(0 To numberOfMonsters - 1) As New clsMONSTERTYPE ' holds monster info
Global Const flailSIZEPX = 14 ' size of flail in pixels

Global intPLAYERS As Integer ' number of current players

Global Const FPS = 60 ' frames per second
Global Const windowX = 700 ' width of frmATTACK window
Global Const windowY = 500 ' height of frmATTACK window
Global Const castleWALLLEFT = 321 ' left wall
Global Const castleWALLRIGHT = 377 ' right wall

' level vars
Public sngMOVESPEED As Single ' speed of monsters
Global arrTOBEMONSTERS() As Integer ' array of to-be monsters
Public intCURRENTMONSTER As Integer ' current monster in the lineup
Public intMONSTERSKILLED As Integer ' number of monsters player has killed
Public intMONSTERSATTACKEDCASTLE As Integer ' number of monsters which have attacked the castle
Public lMONSTERSPAWNCOOLDOWN As Long ' starts at 0, jumps to n when a monster is spawned, -1 each tick, monsters can't spawn if >0
Public bEXIT As Boolean ' somebody won
Public bFORCEEXIT As Boolean ' exited program or stopped server

Sub loadMONSTERINFO()
    loadONEMONSTERINFO greenMonster, "monster0", 9, 25, 1, 1, 2, -1, 1, 0, 2
    loadONEMONSTERINFO blackMonster, "monster1", 9, 25, 2, 2, 5, -1, 1, 1, 2
    loadONEMONSTERINFO bat, "monster2", 10, 11, 2, 1, 3, 150, 1.5, 0, 2
    loadONEMONSTERINFO tree, "monster3", 26, 50, 5, 20, 8, -1, 0.4, 1, 5
    loadONEMONSTERINFO cloud, "monster4", 43, 28, 4, 3, 5, 10, 1, 1, 3
    loadONEMONSTERINFO rabbit, "monster5", 17, 34, 3, 4, 3, -1, 2, 1, 3
    loadONEMONSTERINFO ladybug, "monster6", 13, 7, 1, 4, 2, -1, 2.5, 1, 2
    loadONEMONSTERINFO knightSword, "knight", 21, 51, 5, 10, 20, -1, 0.5, 1, 4
    loadONEMONSTERINFO knightFlail, "knightFlail", 33, 51, 5, 15, 35, -1, 0.5, 1, 6
    loadONEMONSTERINFO knightHorse, "knightHorse", 92, 43, 7, 8, 20, -1, 3, 1, 8
    loadONEMONSTERINFO dragon, "dragon", 91, 53, 10, 50, 200, 200, 0.3, 0, 10
    ' note to self: when adding monsters, change numberOfMonsters
End Sub

Function getMOVESPEED() As Single
    getMOVESPEED = 1 + ((lCURRENTLEVEL * intPLAYERS) / 10)
End Function

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

Sub generateMONSTERS(ByRef lLEVELPOINTS As Long)
    Dim intNEWMONSTER As Integer
    Dim intSTARTINGMONSTER As Integer ' stops infinite loop if there isn't a monster worth 1 point
    Dim intCURRENTMON As Integer
    intCURRENTMON = -1
    
    Do While lLEVELPOINTS > 0
        intNEWMONSTER = Int(Rnd() * numberOfMonsters) ' random monster
        intSTARTINGMONSTER = intNEWMONSTER ' record starting monster
        Do While cmontypeMONSTERINFO(intNEWMONSTER).intPOINTCOST > lLEVELPOINTS ' while monsters have too many points
            intNEWMONSTER = intNEWMONSTER + 1 ' get the next monster
            If intNEWMONSTER = intSTARTINGMONSTER Then ' if back at starting monster
                Exit Do ' not enough points to get any monster
            End If
            If intNEWMONSTER = numberOfMonsters Then ' reached upper bound of monsters
                intNEWMONSTER = 0 ' set to bottom of monsters
            End If
        Loop
        intCURRENTMON = intCURRENTMON + 1 ' one more monster added
        ReDim arrTOBEMONSTERS(0 To intCURRENTMON) ' add spot for the monster
        arrTOBEMONSTERS(intCURRENTMON) = intNEWMONSTER ' set new monster
        lLEVELPOINTS = lLEVELPOINTS - cmontypeMONSTERINFO(intNEWMONSTER).intPOINTCOST ' take away points
    Loop
End Sub
