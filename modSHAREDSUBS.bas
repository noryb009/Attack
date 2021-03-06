Attribute VB_Name = "modSHAREDSUBS"
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

Global Const VERSION = "0.0.0.1e" ' version of the game
Global Const MAXCLIENTS = 19 ' max number of clients

Global Const numberOfMonsters = 12 ' number of types of monsters
Public Enum monsterNames ' names of monsters
    greenMonster
    blackMonster
    bat
    tree
    cloud
    goo
    ladybug
    knightSword
    knightFlail
    knightHorse
    motherBat
    dragon
End Enum

Global lPLAYERCOLOURS(0 To MAXCLIENTS - 1) As Long

Global Const landHEIGHT = 376 ' height of land
Global cmontypeMONSTERINFO(0 To numberOfMonsters - 1) As New clsMONSTERTYPE ' holds monster info
Global Const flailSIZEPX = 14 ' size of flail in pixels
Global Const ticksPerFrame = 6 ' timer ticks per animation frame

Global Const lMONSTERARRAYSIZE As Long = 100
Global Const lFLAILARRAYSIZE As Long = 100

Global intPLAYERS As Integer ' number of current players

Global Const FPS = 60 ' frames per second
Global Const windowX = 700 ' width of frmATTACK window
Global Const windowY = 500 ' height of frmATTACK window
Global Const castleWALLLEFT = 321 ' left wall
Global Const castleWALLRIGHT = 377 ' right wall
Global Const maxLENGTHOFMSGINGAME = 30 ' length of a message to show while in game

' level vars
Public sngMOVESPEED As Single ' speed of monsters
Global arrTOBEMONSTERS() As Integer ' array of to-be monsters
Public lCURRENTMONSTER As Long ' current monster in the lineup
Public lMONSTERSKILLED As Long ' number of monsters player has killed
Public lMONSTERSATTACKEDCASTLE As Long ' number of monsters which have attacked the castle
Public lMONSTERSPAWNCOOLDOWN As Long ' starts at 0, jumps to n when a monster is spawned, -1 each tick, monsters can't spawn if >0
Public bEXIT As Boolean ' somebody won
Public bFORCEEXIT As Boolean ' exited program or stopped server

Global arrMONSTERS(0 To lMONSTERARRAYSIZE - 1) As New clsMONSTER ' monsters
Global arrFLAILS(0 To lFLAILARRAYSIZE - 1) As New clsFLAIL ' flails

' upgrades
Global intFLAILPOWER As Integer ' the attack power of the flails
Global intFLAILGOTHROUGH As Integer ' the number of monsters a flail can go through
Global intFLAILAMOUNT As Integer ' the amount of flails thrown

' castle health
Global lCASTLECURRENTHEALTH As Long ' current castle health
Global lCASTLEMAXHEALTH As Long ' castle max health

Sub loadMONSTERINFO()
    ' monster info
    ' number in enum, image filename, image width, image height, spawn point cost, health,
    '   attack power, Y location (-1 is ground), X speed, money given when hit,
    '   money given when killed[, Y speed]
    loadONEMONSTERINFO greenMonster, "greenMonster", 9, 25, 1, 1, 2, -1, 1, 0, 2
    loadONEMONSTERINFO blackMonster, "blackMonster", 9, 25, 2, 2, 3, -1, 1, 1, 2
    loadONEMONSTERINFO bat, "bat", 10, 11, 2, 1, 3, 150, 1.5, 0, 2, 0.4
    loadONEMONSTERINFO motherBat, "motherBat", 61, 39, 20, 30, 100, 150, 1, 0, 8, 0.2
    loadONEMONSTERINFO tree, "tree", 26, 50, 5, 20, 8, -1, 0.4, 1, 5
    loadONEMONSTERINFO cloud, "cloud", 42, 31, 4, 6, 5, 10, 1, 1, 3, 0.65
    loadONEMONSTERINFO goo, "goo", 24, 11, 5, 6, 5, -1, 2, 1, 3
    loadONEMONSTERINFO ladybug, "ladyBug", 13, 7, 1, 4, 2, -1, 2.5, 1, 2
    loadONEMONSTERINFO knightSword, "knight", 49, 54, 5, 10, 20, -1, 0.5, 1, 4
    loadONEMONSTERINFO knightFlail, "knightFlail", 47, 54, 5, 15, 35, -1, 0.5, 1, 6
    loadONEMONSTERINFO knightHorse, "knightHorse", 77, 64, 7, 8, 20, -1, 3, 1, 8
    loadONEMONSTERINFO dragon, "dragon", 90, 57, 50, 50, 200, 200, 1, 0, 10
    ' note to self: when adding monsters, change numberOfMonsters
End Sub

Sub loadPLAYERCOLOURS()
    lPLAYERCOLOURS(0) = RGB(96, 0, 0) ' red
    lPLAYERCOLOURS(1) = RGB(174, 174, 174) ' grey
    lPLAYERCOLOURS(2) = RGB(255, 128, 128) ' dark pink
    lPLAYERCOLOURS(3) = RGB(0, 0, 128) ' dark blue
    lPLAYERCOLOURS(4) = RGB(30, 255, 30) ' light green
    lPLAYERCOLOURS(5) = RGB(128, 0, 255) ' dark purple
    lPLAYERCOLOURS(6) = RGB(216, 127, 0) ' yellow
    lPLAYERCOLOURS(7) = RGB(216, 127, 0) ' orange
    lPLAYERCOLOURS(8) = RGB(128, 64, 0) ' brown
    lPLAYERCOLOURS(9) = RGB(255, 0, 128) ' pink
    lPLAYERCOLOURS(10) = RGB(109, 109, 109) ' dark grey
    lPLAYERCOLOURS(11) = RGB(191, 191, 0) ' dark yellow
    lPLAYERCOLOURS(12) = RGB(128, 0, 128) ' dark purple
    lPLAYERCOLOURS(13) = RGB(2, 2, 2) ' black
    lPLAYERCOLOURS(14) = RGB(255, 0, 255) ' light purple
    lPLAYERCOLOURS(15) = RGB(0, 255, 255) ' cyan
    lPLAYERCOLOURS(16) = RGB(0, 128, 0) ' dark green
    lPLAYERCOLOURS(17) = RGB(0, 255, 128) ' teal
    lPLAYERCOLOURS(18) = RGB(0, 64, 64) ' dark blue
End Sub

Function getMOVESPEED() As Single ' get the movement speed (used by server and online client)
    getMOVESPEED = 1 + ((lCURRENTLEVEL * intPLAYERS) / 10) ' return formula to get move speed
End Function

Function getMONSTERSLEFT() As String ' returns the monsters left in the level
    getMONSTERSLEFT = CStr((UBound(arrTOBEMONSTERS) + 1) - lMONSTERSKILLED - lMONSTERSATTACKEDCASTLE) ' calculate the monsters left
End Function

Sub doEVENTSANDSLEEP(lMILLISECONDS As Long)
    Dim nC As Long
    nC = 0
    Do While (nC < lMILLISECONDS \ 100) ' for each 10th of a second in lMILLISECONDS
        DoEvents ' do any events that need to be done
        Sleep 100 ' save CPU and sleep
        nC = nC + 1 ' next 10th of a second
    Loop
    Sleep lMILLISECONDS Mod 100 ' sleep for rest of time
End Sub

Function addZEROIFNOTZERO(lVALUE As Long) As String
    If lVALUE = 0 Then
        addZEROIFNOTZERO = CStr(lVALUE)
    Else
        addZEROIFNOTZERO = CStr(lVALUE) & "0"
    End If
End Function

Function safeADDLONG(lNUMBER1 As Long, lNUMBER2 As Long) As Long ' add longs without overflows
    If lNUMBER2 = 0 Then ' if not adding anything
        safeADDLONG = lNUMBER1 ' return first number
        Exit Function ' exit
    End If
    Dim dblOUTPUT As Double ' temp double (can hold more then longs, so no overflow)
    dblOUTPUT = lNUMBER1 ' store first number
    dblOUTPUT = dblOUTPUT + lNUMBER2 ' add second number
    If dblOUTPUT < 2147483647 Then ' if number won't overflow a long
        safeADDLONG = dblOUTPUT ' return number
    Else ' dblOUTPUT would overflow long
        safeADDLONG = 2147483647 ' return long max
    End If
End Function

Sub generateMONSTERS(ByRef lLEVELPOINTS As Long) ' generate monsters
    Dim intNEWMONSTER As Integer ' random new monster
    Dim intSTARTINGMONSTER As Integer ' stops infinite loop if there isn't a monster worth 1 point
    Dim intCURRENTMON As Integer ' number of current monster
    intCURRENTMON = -1 ' starts with -1 monsters, adds one, and starts at 0
    
    Do While lLEVELPOINTS > 0 ' do while you still have points
        If numberOfMonsters < (lCURRENTLEVEL + 2) Then ' if you have unlocked all the monsters
            intNEWMONSTER = Int(Rnd() * numberOfMonsters) ' random monster
        Else ' you haven't unlocked all the monsters yet
            intNEWMONSTER = Int(Rnd() * (lCURRENTLEVEL + 2)) ' random monster from the monsters unlocked so far
        End If
        
        intSTARTINGMONSTER = intNEWMONSTER ' record starting monster
        
        Do While cmontypeMONSTERINFO(intNEWMONSTER).intPOINTCOST > lLEVELPOINTS ' while monsters have too many points
            intNEWMONSTER = intNEWMONSTER + 1 ' get the next monster
            If intNEWMONSTER = intSTARTINGMONSTER Then ' if back at starting monster
                Exit Do ' not enough points to get any monster
            End If
            If intNEWMONSTER = numberOfMonsters Or intNEWMONSTER = (lCURRENTLEVEL + 2) Then ' reached upper bound of monsters, or reached max monster for current level
                intNEWMONSTER = 0 ' set to bottom of monsters
            End If
        Loop
        intCURRENTMON = intCURRENTMON + 1 ' one more monster added
        ReDim Preserve arrTOBEMONSTERS(0 To intCURRENTMON) ' add spot for the monster
        arrTOBEMONSTERS(intCURRENTMON) = intNEWMONSTER ' set new monster
        lLEVELPOINTS = lLEVELPOINTS - cmontypeMONSTERINFO(intNEWMONSTER).intPOINTCOST ' take away points
    Loop
End Sub
