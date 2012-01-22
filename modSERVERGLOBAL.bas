Attribute VB_Name = "modGLOBAL"
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

Global Const SERVER = True ' is server
Global Const programNAME = "Attack Server"
Global Const onlineMODE = False ' used for being able to share subs with client

Global cCLIENTS(0 To MAXCLIENTS - 1) As New clsCONNECTION ' client connections
Global cCLIENTINFO(0 To MAXCLIENTS - 1) As New clsCLIENTINFO ' client info
Global lWINNEROFROUND As Long ' winner of last round
Global lMONEY As Long ' money

Global lCURRENTLEVEL As Long ' level server is on

Global bPLAYING As Boolean ' if game running

' get the number of ticks so far
Public Declare Function QueryPerformanceCounter Lib "kernel32" ( _
    lpPerformanceCount As Currency _
) As Long

' frequency of ticks
Public Declare Function QueryPerformanceFrequency Lib "kernel32" ( _
    lpFrequency As Currency _
) As Long

' sleep, don't use CPU for a bit
Public Declare Sub Sleep Lib "kernel32" ( _
    ByVal dwMilliseconds As Long _
)

Public Sub log(strNEWLINE As String) ' add to log
    frmSERVER.txtLOG.Text = strNEWLINE & vbCrLf & frmSERVER.txtLOG.Text ' add the new text, a new line, then the old text
End Sub

Public Sub moveEVERYTHING() ' move all the monsters and flails
    Dim nC As Long
    
    ' move monsters
    Dim lHEALTHBEFOREMOVEMON As Long
    lHEALTHBEFOREMOVEMON = lCASTLECURRENTHEALTH ' store health before monsters attack
    nC = 0
    Do While nC < lMONSTERARRAYSIZE ' for each monster
        If arrMONSTERS(nC).bACTIVE = True Then ' if monster is enabled
            arrMONSTERS(nC).moveMONSTER ' move the monster
            
            If arrMONSTERS(nC).bACTIVE = False Then ' if monster got disabled (attacked castle or out of the screen)
                broadcastMONSTER nC ' broadcast monster state
                broadcast "monstersLeft", getMONSTERSLEFT ' sync number of monsters left
            End If
        End If
        nC = nC + 1 ' next monster
    Loop
    
    If lHEALTHBEFOREMOVEMON <> lCASTLECURRENTHEALTH Then ' if health changed
        If lCASTLECURRENTHEALTH <= 0 Then ' if out of health
            bEXIT = True ' exit, you lost
            broadcast "health", "0" ' send health to clients
        Else
            broadcast "health", CStr(lCASTLECURRENTHEALTH) ' send health to clients
        End If
    End If
    
    ' move flails
    Dim lADDTOSCORE As Long ' amount to add to flail owner's score
    nC = 0
    Do While nC < lFLAILARRAYSIZE ' for each flail
        If arrFLAILS(nC).bACTIVE = True Then ' if flail is on screen
            lADDTOSCORE = arrFLAILS(nC).moveFLAIL ' move flail and get score to add
            If lADDTOSCORE <> 0 Then ' if flail hit something
                cCLIENTINFO(arrFLAILS(nC).lOWNER).lLEVELSCORE = safeADDLONG(cCLIENTINFO(arrFLAILS(nC).lOWNER).lLEVELSCORE, lADDTOSCORE) ' add score
                'cCLIENTS(arrFLAILS(nC).lOWNER).sendString "moneyLevel", CStr(cCLIENTINFO(arrFLAILS(nC).lOWNER).lLEVELSCORE) ' alert owner of new current score
                broadcast "playerScore", CStr(arrFLAILS(nC).lOWNER) & "\" & CStr(cCLIENTINFO(arrFLAILS(nC).lOWNER).lLEVELSCORE)
            End If
            If arrFLAILS(nC).bACTIVE = False Then ' if flail got deactivated
                broadcastFLAIL nC, False ' broadcast that flail is no longer active
            End If
        End If
        nC = nC + 1 ' next flail
    Loop
    
    ' spawn monsters
    If lMONSTERSPAWNCOOLDOWN = 0 And lCURRENTMONSTER <= UBound(arrTOBEMONSTERS) Then ' if monster hasn't been spawned for a while and monsters still waiting
        Dim bSPAWN As Boolean
        bSPAWN = False ' default: don't spawn
        If lCURRENTMONSTER <= lMONSTERSKILLED + lMONSTERSATTACKEDCASTLE + (lCURRENTLEVEL \ 3) + intPLAYERS Then ' force if less then (level/3) monsters on screen
            bSPAWN = True ' spawn
        ElseIf Int(Rnd() * 150) < lCURRENTLEVEL * intPLAYERS + intPLAYERS Then ' randomly
            bSPAWN = True ' spawn
        End If
    
        If bSPAWN = True Then ' if monster going to be spawned
            spawnMONSTER ' spawn the monster
            lMONSTERSPAWNCOOLDOWN = 20 ' set cool down time
        End If
    Else ' needs to cool down more
        lMONSTERSPAWNCOOLDOWN = lMONSTERSPAWNCOOLDOWN - 1 ' count down cooldown time
    End If
    
    If lMONSTERSKILLED + lMONSTERSATTACKEDCASTLE > UBound(arrTOBEMONSTERS) Then ' if you defeated all the monsters on this level
        bEXIT = True ' exit
    End If
End Sub

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

Public Sub startGAME() ' start the game
    If lCASTLECURRENTHEALTH = 0 Then ' if you don't have any health
        lCASTLECURRENTHEALTH = 1 ' give yourself some
        broadcast "health", "1" ' broadcast health
    End If
    
    If bPLAYING = True Then ' if already playing a game
        Exit Sub ' exit
    End If
    
    bPLAYING = True ' you are playing the game
    
    Dim nC As Integer
    
    ' reset level vars
    bEXIT = False
    bFORCEEXIT = False
    lMONSTERSKILLED = 0
    lMONSTERSATTACKEDCASTLE = 0
    lCURRENTMONSTER = 0
    lMONSTERSPAWNCOOLDOWN = 0
    ReDim arrTOBEMONSTERS(0 To 0)
    sngMOVESPEED = getMOVESPEED
    
    ' clear monsters array
    nC = 0
    Do While nC < lMONSTERARRAYSIZE ' for each monster spot
        arrMONSTERS(nC).bACTIVE = False ' monster is not active
        nC = nC + 1 ' next monster spot
    Loop
    
    ' clear flail array
    nC = 0
    Do While nC < lMONSTERARRAYSIZE ' for each flail spot
        arrFLAILS(nC).bACTIVE = False ' flail is not active
        nC = nC + 1 ' next flail spot
    Loop
    
    broadcast "disableReadyButton", "" '  disable clients' ready button
    
    log "Starting countdown..." ' log that you are the countdown game
    broadcast "chat", "Game starting in 5..." ' broadcast time left
    doEVENTSANDSLEEP 1000 ' wait for 1 second
    broadcast "chat", "4..." ' broadcast time left
    doEVENTSANDSLEEP 1000 ' wait for 1 second
    broadcast "chat", "3..." ' broadcast time left
    doEVENTSANDSLEEP 1000 ' wait for 1 second
    broadcast "chat", "2..." ' broadcast time left
    doEVENTSANDSLEEP 1000 ' wait for 1 second
    broadcast "chat", "1..." ' broadcast time left
    doEVENTSANDSLEEP 1000 ' wait for 1 second
    broadcast "chat", "0..." ' broadcast time left
    
    If bFORCEEXIT = True Then ' if program exited or everyone left during count down
        bPLAYING = False ' you are not playing the game
        Exit Sub ' exit
    End If
    
    log "Starting game..." ' log that you are starting game
    broadcast "game", "start" ' broadcast to clients that the game is starting
    
    ' reset ready status
    nC = 0
    Do While nC < MAXCLIENTS ' for each client
        cCLIENTINFO(nC).bREADY = False ' not ready
        nC = nC + 1 ' next client
    Loop
    
    ' generate level monsters
    generateMONSTERS (lCURRENTLEVEL * 20) + 30 ' generate monsters
    broadcast "monstersLeft", getMONSTERSLEFT ' sync number of monsters left
    
    frmSERVER.timerSYNC.Enabled = True ' sync with clients
    
    doEVENTSANDSLEEP 1000 ' give clients a little time to get ready
    
    Dim currSTARTTIME As Currency ' starting time
    Dim currCURRENTTIME As Currency ' current time
    'Dim currFREQUENCY As Currency ' frame frequency
    Dim dblTIMEBETWEENFRAMES As Double ' time between frames
    
    QueryPerformanceFrequency currCURRENTTIME ' currFREQUENCY ' get the frequency of ticks
    dblTIMEBETWEENFRAMES = currCURRENTTIME / FPS ' currFREQUENCY / FPS ' get time between frames needed to reach FPS
    
    Do While bEXIT = False And bFORCEEXIT = False ' if not exiting yet
        QueryPerformanceCounter currCURRENTTIME ' get current time
        If currCURRENTTIME >= currSTARTTIME + dblTIMEBETWEENFRAMES Then ' if start time + time between frame = current time, then time for the next frame
            QueryPerformanceCounter currSTARTTIME ' store current time as new start time
            moveEVERYTHING ' move everything
        Else
            Sleep 1
        End If
        DoEvents ' do any events needed to be done
    Loop
    
    frmSERVER.timerSYNC.Enabled = False ' don't sync with clients
    
    log "Stopping game..." ' log that the server is stopping
    bPLAYING = False ' you stopped playing the game
    
    If bFORCEEXIT = False Then ' if program not closing
        ' allow shared win/loose code
        Dim sngMONEYMULTIPLYER As Single
        Dim strWINLOOSE As String ' if you won/lost
        
        If lCASTLECURRENTHEALTH <= 0 Then ' if no more health, you lost
            lCASTLECURRENTHEALTH = 1 ' reset current health, enough to continue if round winner doesn't buy health
            sngMONEYMULTIPLYER = 0.5 ' half money
            strWINLOOSE = "Loose" ' you lost
            If lCURRENTLEVEL > 0 Then ' if not at easiest level
                lCURRENTLEVEL = lCURRENTLEVEL - 1 ' go back a level
            End If
        Else ' health left, you won
            sngMONEYMULTIPLYER = 1 ' all of money
            strWINLOOSE = "Win" ' you won
            lCURRENTLEVEL = lCURRENTLEVEL + 1 ' next level
        End If
        
        Dim lHIGHESTSCORE As Long ' highest score found so far
        lWINNEROFROUND = -1 ' nobody won yet
        
        ' add score and find winner
        nC = 0
        Do While nC < MAXCLIENTS ' for each client
            If cCLIENTS(nC).connected = True Then ' if connected
                lMONEY = safeADDLONG(lMONEY, CLng(cCLIENTINFO(nC).lLEVELSCORE / sngMONEYMULTIPLYER)) ' add (half if you lost) client's money to total money
                If lHIGHESTSCORE < cCLIENTINFO(nC).lLEVELSCORE Or (lWINNEROFROUND = -1 And cCLIENTINFO(nC).lLEVELSCORE > 0) Then ' if highest score so far or first client
                    lWINNEROFROUND = nC ' you are the current winner of the round
                    lHIGHESTSCORE = cCLIENTINFO(nC).lLEVELSCORE ' remember your score
                End If
            End If
            nC = nC + 1 ' next client
        Loop
        
        broadcast "health", CStr(lCASTLECURRENTHEALTH) ' broadcast health
        broadcast "moneyTotal", CStr(lMONEY) ' broadcast money
        broadcast "nextLevel", CStr(lCURRENTLEVEL) ' broadcast current level
        
        nC = 0
        
        Do While nC < MAXCLIENTS ' for each client
            If cCLIENTS(nC).connected = True Then ' if client is connected
                If nC = lWINNEROFROUND Then ' if client won
                    cCLIENTS(nC).sendString "game", "stop" & strWINLOOSE & "Highscore" ' alert user that they won the round
                Else
                    cCLIENTS(nC).sendString "game", "stop" & strWINLOOSE ' alert user that the game has ended
                End If
            End If
            nC = nC + 1 ' next client
        Loop
        
        If lWINNEROFROUND <> -1 Then ' if found a winner
            broadcast "chat", formatCHATMSG(cCLIENTINFO(lWINNEROFROUND).strNAME, lPLAYERCOLOURS(lWINNEROFROUND)) & formatCHATMSG(" won the round with a score of " & cCLIENTINFO(lWINNEROFROUND).lLEVELSCORE & "0!") ' broadcast winner of round
        End If
        
        ' reset client vars
        nC = 0
        Do While nC < MAXCLIENTS ' for each client
            If cCLIENTS(nC).connected = True Then ' if connected
                cCLIENTINFO(nC).afterLevelReset ' reset client vars
            End If
            nC = nC + 1 ' next client
        Loop
    End If
    
    'frmSERVER.timerGAME.Enabled = True ' start game timer
End Sub

Sub broadcast(strCOMMAND As String, strTOSEND As String) ' send a command to all the clients
    Dim nC As Integer
    nC = 0
    Do While nC < MAXCLIENTS ' for each client
        If cCLIENTS(nC).connected = True Then ' if connected
            cCLIENTS(nC).sendString strCOMMAND, strTOSEND ' send string to client
            'DoEvents ' do events (including sending this request)
        End If
        nC = nC + 1 ' next client
    Loop
End Sub

Function getMONSTERINFO(lMONSTERNUMBER As Long) As String ' get monster info into a string
    ' return all the monster info in a string
    getMONSTERINFO = CStr(lMONSTERNUMBER) & "~" & _
    CStr(arrMONSTERS(lMONSTERNUMBER).bACTIVE) & "~" & _
    CStr(arrMONSTERS(lMONSTERNUMBER).intTYPE) & "~" & _
    CStr(arrMONSTERS(lMONSTERNUMBER).sngX) & "~" & _
    CStr(arrMONSTERS(lMONSTERNUMBER).sngY) & "~" & _
    CStr(arrMONSTERS(lMONSTERNUMBER).sngMOVINGH) & "~" & _
    CStr(arrMONSTERS(lMONSTERNUMBER).intHEALTH)
End Function

Sub syncMONSTERS() ' send all of the monsters to clients
    Dim nC As Long
    nC = 0
    Dim strTOSEND As String ' string going to be sent to clients
    Do While nC < lMONSTERARRAYSIZE ' for each monster
        If arrMONSTERS(nC).bACTIVE = True Then ' only sync if active
            If strTOSEND <> "" Then ' if not first
                strTOSEND = strTOSEND & "\" ' add separator
            End If
            strTOSEND = strTOSEND & getMONSTERINFO(nC) ' add current monster's info
        End If
        nC = nC + 1 ' next monster
    Loop
    broadcast "syncMon", strTOSEND ' send all the monster info to clients
End Sub

Sub broadcastMONSTER(lMONSTERNUMBER As Long) ' send one monster's info to clients
    broadcast "updateMon", getMONSTERINFO(lMONSTERNUMBER) ' send monster data to clients
End Sub

Function getFLAILINFO(lFLAILNUMBER As Long, bCLEARGOTHROUGH As Boolean) As String ' format flail info into a string
    ' return flail info
    getFLAILINFO = _
    CStr(lFLAILNUMBER) & "~" & _
    CStr(arrFLAILS(lFLAILNUMBER).bACTIVE) & "~" & _
    CStr(arrFLAILS(lFLAILNUMBER).lOWNER) & "~" & _
    CStr(arrFLAILS(lFLAILNUMBER).sngX) & "~" & _
    CStr(arrFLAILS(lFLAILNUMBER).sngY) & "~" & _
    CStr(arrFLAILS(lFLAILNUMBER).sngMOVINGV) & "~" & _
    CStr(arrFLAILS(lFLAILNUMBER).sngMOVINGH) & "~" & _
    CStr(arrFLAILS(lFLAILNUMBER).intGOTHROUGH) & "~" & _
    CStr(bCLEARGOTHROUGH)
End Function

Sub syncFLAILS() ' send all flails to clients
    Dim nC As Long
    nC = 0
    Dim strTOSEND As String ' string going to be send to clients
    Do While nC < lMONSTERARRAYSIZE ' for each flail
        If arrFLAILS(nC).bACTIVE = True Then ' only sync if active
            If strTOSEND <> "" Then ' if not first
                strTOSEND = strTOSEND & "\" ' add separator
            End If
            strTOSEND = strTOSEND & getFLAILINFO(nC, False) ' add flail info
        End If
        nC = nC + 1 ' next flail
    Loop
    broadcast "syncFla", strTOSEND ' send all the flails' info to clients
End Sub

Sub broadcastFLAIL(lFLAILNUMBER As Long, bCLEARGOTHROUGH As Boolean) ' send one flail to clients
    broadcast "updateFlail", getFLAILINFO(lFLAILNUMBER, bCLEARGOTHROUGH) ' send flail info
End Sub

Function formatCHATMSG(ByVal strMESSAGE As String, Optional lCOLOUR As Long = vbBlack, Optional lBOLD As Long = -1) As String
    strMESSAGE = Replace(strMESSAGE, "&", "&amp;") ' escape &
    strMESSAGE = Replace(strMESSAGE, "\", "&bslash;") ' escape \
    strMESSAGE = Replace(strMESSAGE, "~", "&tide;") ' escape ~
    
    If (lBOLD = -1 And lCOLOUR <> vbBlack) Or lBOLD = 1 Then ' if force bold or (detect bold mode and message is coloured)
        formatCHATMSG = CStr(lCOLOUR) & "\" & CStr(True) & "\" & strMESSAGE & "~" ' colour\bold\message~colour\message~...
    Else ' if force not bold or (detect bold mode and message is not coloured)
        formatCHATMSG = CStr(lCOLOUR) & "\" & CStr(False) & "\" & strMESSAGE & "~" ' colour\bold\message~colour\message~...
    End If
End Function

Public Sub spawnMONSTER() ' spawn a monster
    Dim nC As Long
    nC = 0
    Do While nC < lMONSTERARRAYSIZE ' for each monster
        If arrMONSTERS(nC).bACTIVE = False Then ' if not active
            arrMONSTERS(nC).bACTIVE = True ' monster is now active
            arrMONSTERS(nC).intTYPE = arrTOBEMONSTERS(lCURRENTMONSTER) ' next monster
            
            arrMONSTERS(nC).sngX = Int(Rnd() * 2) ' randomize starting side
            If arrMONSTERS(nC).sngX = 0 Then ' if on left side
                arrMONSTERS(nC).sngX = 0 - cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).lWIDTH ' start at left side
                arrMONSTERS(nC).sngMOVINGH = 1 * cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).sngXSPEED ' go right
            Else ' if on right side
                arrMONSTERS(nC).sngX = windowX ' start of right side
                arrMONSTERS(nC).sngMOVINGH = -1 * cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).sngXSPEED ' go left
            End If
            arrMONSTERS(nC).sngY = cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).intSTARTINGY ' start at starting Y location
            arrMONSTERS(nC).sngMOVINGV = cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).sngYSPEED ' set vertical going down speed
            arrMONSTERS(nC).intHEALTH = cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).intMAXHEALTH ' set to max health
            
            lCURRENTMONSTER = safeADDLONG(lCURRENTMONSTER, 1) ' one more monster placed
            broadcastMONSTER nC ' broadcast new monster
            Exit Do ' found empty monster spot, continue
        End If
        nC = nC + 1 ' next monster spot
    Loop
End Sub

' load monster info
Sub loadONEMONSTERINFO(intNUMBER As Integer, imageNAME As String, lIMAGEWIDTH As Long, lIMAGEHEIGHT As Long, intPOINTCOST As Integer, intHEALTH As Integer, intATTACKPOWER As Integer, intSTARTINGY As Integer, sngXSPEED As Single, intMONEYONHIT As Integer, intMONEYONKILL As Integer, Optional sngYSPEED As Single = 0)
    cmontypeMONSTERINFO(intNUMBER).intPOINTCOST = intPOINTCOST ' load point cost
    cmontypeMONSTERINFO(intNUMBER).intMAXHEALTH = intHEALTH ' load health
    cmontypeMONSTERINFO(intNUMBER).intATTACKPOWER = intATTACKPOWER ' load attack power
    cmontypeMONSTERINFO(intNUMBER).sngXSPEED = sngXSPEED ' load vertical speed
    cmontypeMONSTERINFO(intNUMBER).sngYSPEED = sngYSPEED ' load horizontal speed
    If intSTARTINGY = -1 Then ' default: ground
        cmontypeMONSTERINFO(intNUMBER).intSTARTINGY = landHEIGHT - lIMAGEHEIGHT ' start at land - image height, so feet are on land
    Else ' not on ground
        cmontypeMONSTERINFO(intNUMBER).intSTARTINGY = intSTARTINGY ' load starting Y
    End If
    cmontypeMONSTERINFO(intNUMBER).intMONEYADDEDHIT = intMONEYONHIT ' load money on hit
    cmontypeMONSTERINFO(intNUMBER).intMONEYADDEDKILL = intMONEYONKILL ' load money on kill
    cmontypeMONSTERINFO(intNUMBER).lWIDTH = lIMAGEWIDTH ' load image width
    cmontypeMONSTERINFO(intNUMBER).lHEIGHT = lIMAGEHEIGHT ' load image height
End Sub

Sub Main() ' program init
    Randomize ' randomize random numbers
    
    loadMONSTERINFO ' load monster info into cmontypeMONSTERINFO()
    loadPLAYERCOLOURS ' load player colour info into playerCOLOURS()
    
    Dim nC As Integer
    nC = 0
    Do While nC < MAXCLIENTS ' for each client spot
        cCLIENTS(nC).arrayID = nC ' set client ID in the array
        nC = nC + 1 ' next client spot
    Loop
    
    ' reset monster array
    nC = 0
    Do While nC < lMONSTERARRAYSIZE ' for each monster spot
        nC = nC + 1 ' next monster spot
    Loop
    
    ' reset flail array
    nC = 0
    Do While nC < lFLAILARRAYSIZE ' for each flail spot
        nC = nC + 1 ' next flail spot
    Loop
    
    ReDim arrTOBEMONSTERS(0 To 0) ' resize array to have no current monsters
    
    frmSERVER.Show ' show the server form
End Sub
