Attribute VB_Name = "modGLOBAL"
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

Global Const VERSION = "0.0.0.1s"
Global Const SERVER = True
Global Const onlineMODE = False
Global Const MAXCLIENTS = 4

'Global bHASINFO As Boolean

Global lPORT As Long

Global cCLIENTS(0 To MAXCLIENTS - 1) As New clsCONNECTION
Global cCLIENTINFO(0 To MAXCLIENTS - 1) As New clsCLIENTINFO
Global lWINNEROFROUND As Long
Global lMONEY As Long

Global lCURRENTLEVEL As Long
Global lLEVELMONEY As Long

' upgrades
Global intFLAILPOWER As Integer ' the attack power of the flails
Global intFLAILGOTHROUGH As Integer ' the number of monsters a flail can go through
Global intFLAILAMOUNT As Integer ' the amount of flails thrown

' castle health
Global lCASTLECURRENTHEALTH As Long
Global lCASTLEMAXHEALTH As Long

Global arrMONSTERS() As clsMONSTER
Global arrFLAILS() As clsFLAIL

Public Declare Function QueryPerformanceCounter Lib "kernel32" ( _
    lpPerformanceCount As Currency _
) As Long

Public Declare Function QueryPerformanceFrequency Lib "kernel32" ( _
    lpFrequency As Currency _
) As Long

Public Declare Sub Sleep Lib "kernel32" ( _
    ByVal dwMilliseconds As Long _
)

Public Sub log(strNEWLINE As String)
    frmSERVER.txtLOG.Text = strNEWLINE & vbCrLf & frmSERVER.txtLOG.Text
End Sub

Public Sub moveEVERYTHING()
    Dim nC As Long
    
    ' spawn monsters
    If lMONSTERSPAWNCOOLDOWN = 0 Then
        Dim bSPAWN As Boolean
        bSPAWN = False
        If intCURRENTMONSTER <= intMONSTERSKILLED + intMONSTERSATTACKEDCASTLE + (lCURRENTLEVEL \ 3) + intPLAYERS Then ' force if nobody on screen
            bSPAWN = True
        ElseIf Int(Rnd() * 200) < lCURRENTLEVEL * intPLAYERS + intPLAYERS And intCURRENTMONSTER <= UBound(arrTOBEMONSTERS) Then ' randomly if some monsters are waiting
            bSPAWN = True
        End If
    
        If bSPAWN = True Then
            spawnMONSTER
            lMONSTERSPAWNCOOLDOWN = 20
        End If
    Else
        lMONSTERSPAWNCOOLDOWN = lMONSTERSPAWNCOOLDOWN - 1 ' count down
    End If
    
    ' move monsters
    nC = 0
    Do While nC <= UBound(arrMONSTERS)
        If arrMONSTERS(nC).bACTIVE = True Then
            arrMONSTERS(nC).moveMONSTER
            If (arrMONSTERS(nC).sngMOVINGH < 0 And arrMONSTERS(nC).sngX + cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).lWIDTH < 0) Or (arrMONSTERS(nC).sngMOVINGH > 0 And arrMONSTERS(nC).sngX > windowX) Then
                arrMONSTERS(nC).bACTIVE = False
                broadcastMONSTER nC
            ElseIf (arrMONSTERS(nC).sngMOVINGH > 0 And arrMONSTERS(nC).sngX + cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).lWIDTH > castleWALLLEFT) Or (arrMONSTERS(nC).sngMOVINGH < 0 And arrMONSTERS(nC).sngX < castleWALLRIGHT) Then 'attack
                lCASTLECURRENTHEALTH = lCASTLECURRENTHEALTH - cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).intATTACKPOWER
                intMONSTERSATTACKEDCASTLE = intMONSTERSATTACKEDCASTLE + 1
                arrMONSTERS(nC).bACTIVE = False
                If lCASTLECURRENTHEALTH <= 0 Then
                    bEXIT = True
                    broadcast "health", "0"
                Else
                    broadcast "health", CStr(lCASTLECURRENTHEALTH)
                End If
                broadcastMONSTER nC
            End If
        End If
        nC = nC + 1
    Loop
    
    ' move flails
    nC = 0
    Do While nC <= UBound(arrFLAILS)
        If arrFLAILS(nC).bACTIVE = True Then
            Dim intNEWX As Integer
            Dim intNEWY As Integer
            Dim intDELETEMONSTER As Integer
            intDELETEMONSTER = -1 ' don't delete a monster
            intNEWX = Int(arrFLAILS(nC).sngX) + arrFLAILS(nC).sngMOVINGH
            intNEWY = Int(arrFLAILS(nC).sngY) + arrFLAILS(nC).sngMOVINGV
            
            'gravity
            arrFLAILS(nC).sngMOVINGV = arrFLAILS(nC).sngMOVINGV + 0.5
            If arrFLAILS(nC).sngMOVINGH < 0 Then
                arrFLAILS(nC).sngMOVINGH = arrFLAILS(nC).sngMOVINGH + 0.1
            ElseIf arrFLAILS(nC).sngMOVINGH > 0 Then
                arrFLAILS(nC).sngMOVINGH = arrFLAILS(nC).sngMOVINGH - 0.1
            End If
            
            'check if hitting monster
            Dim nCMONSTERS As Integer
            nCMONSTERS = 0
            Do While nCMONSTERS <= UBound(arrMONSTERS)
                If arrMONSTERS(nCMONSTERS).bACTIVE = True Then
                    If (arrFLAILS(nC).sngY < arrMONSTERS(nCMONSTERS).sngY + cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).lHEIGHT And intNEWY + cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).lHEIGHT > arrMONSTERS(nCMONSTERS).sngY) Or _
                    (intNEWY < arrMONSTERS(nCMONSTERS).sngY + cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).lHEIGHT And arrFLAILS(nC).sngY + flailSIZEPX > arrMONSTERS(nCMONSTERS).sngY) Then
                        'If arrFLAILS(nC).sngx + picFLAIL.Width < arrMONSTERS(nCMONSTERS).sngx And intNEWX + picFLAIL.Width > arrMONSTERS(nCMONSTERS).sngx Then
                        If (arrFLAILS(nC).sngX < arrMONSTERS(nCMONSTERS).sngX + cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).lWIDTH And intNEWX + flailSIZEPX > arrMONSTERS(nCMONSTERS).sngX) Or _
                        (intNEWX < arrMONSTERS(nCMONSTERS).sngX + cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).lWIDTH And arrFLAILS(nC).sngX + flailSIZEPX > arrMONSTERS(nCMONSTERS).sngX) Then
                            If arrFLAILS(nC).didGOTHROUGH(nCMONSTERS) = False Then
                                intDELETEMONSTER = nCMONSTERS
                            End If
                        End If
                    End If
                End If
                nCMONSTERS = nCMONSTERS + 1
            Loop
            
            If intDELETEMONSTER <> -1 Then ' delete a monster
                arrMONSTERS(intDELETEMONSTER).intHEALTH = arrMONSTERS(intDELETEMONSTER).intHEALTH - intFLAILPOWER
                If arrMONSTERS(intDELETEMONSTER).intHEALTH < 1 Then
                    arrMONSTERS(intDELETEMONSTER).bACTIVE = False
                    intMONSTERSKILLED = intMONSTERSKILLED + 1
                    cCLIENTINFO(arrFLAILS(nC).lOWNER).lLEVELSCORE = safeADDLONG(cCLIENTINFO(arrFLAILS(nC).lOWNER).lLEVELSCORE, cmontypeMONSTERINFO(arrMONSTERS(intDELETEMONSTER).intTYPE).intMONEYADDEDKILL)
                Else
                    If intFLAILGOTHROUGH > 1 Then arrFLAILS(nC).addGOTHROUGH intDELETEMONSTER
                    cCLIENTINFO(arrFLAILS(nC).lOWNER).lLEVELSCORE = safeADDLONG(cCLIENTINFO(arrFLAILS(nC).lOWNER).lLEVELSCORE, cmontypeMONSTERINFO(arrMONSTERS(intDELETEMONSTER).intTYPE).intMONEYADDEDHIT)
                End If
                
                ' update user
                cCLIENTS(arrFLAILS(nC).lOWNER).sendString "moneyLevel", CStr(cCLIENTINFO(arrFLAILS(nC).lOWNER).lLEVELSCORE)
                
                If arrFLAILS(nC).intGOTHROUGH > 1 Then
                    arrFLAILS(nC).intGOTHROUGH = arrFLAILS(nC).intGOTHROUGH - 1
                Else
                    arrFLAILS(nC).bACTIVE = False
                End If
                broadcastFLAIL nC, False
            End If
            
            arrFLAILS(nC).sngX = intNEWX
            arrFLAILS(nC).sngY = intNEWY
            
            If arrFLAILS(nC).sngX + flailSIZEPX < 0 Or arrFLAILS(nC).sngX > windowX Or arrFLAILS(nC).sngY < -1000 Or arrFLAILS(nC).sngY > windowY - 50 - flailSIZEPX Then
                arrFLAILS(nC).bACTIVE = False
                broadcastFLAIL nC, False
            End If
        End If
        nC = nC + 1
    Loop
    
    If intMONSTERSKILLED + intMONSTERSATTACKEDCASTLE > UBound(arrTOBEMONSTERS) Then
        bEXIT = True
    End If
End Sub

Public Sub startGAME()
    If lCASTLECURRENTHEALTH = 0 Then
        log "Starting game, but you don't have any health."
        Exit Sub
    End If
    
    log "Starting game..."
    lLEVELMONEY = 0
    broadcast "game", "start" ' broadcast to clients that the game is starting
    
    ' reset ready status
    Dim nC As Integer
    nC = 0
    Do While nC < MAXCLIENTS
        cCLIENTINFO(nC).bREADY = False ' not ready
        nC = nC + 1
    Loop
    
    ' reset level vars
    bEXIT = False
    bFORCEEXIT = False
    lLEVELMONEY = 0
    intMONSTERSKILLED = 0
    intMONSTERSATTACKEDCASTLE = 0
    intCURRENTMONSTER = 0
    lMONSTERSPAWNCOOLDOWN = 0
    ReDim arrTOBEMONSTERS(0 To 0)
    sngMOVESPEED = getMOVESPEED
    
    ReDim arrMONSTERS(0 To 99)
    nC = 0
    Do While nC <= UBound(arrMONSTERS)
        Set arrMONSTERS(nC) = New clsMONSTER
        arrMONSTERS(nC).bACTIVE = False
        nC = nC + 1
    Loop
    
    ReDim arrFLAILS(0 To 99)
    nC = 0
    Do While nC <= UBound(arrFLAILS)
        Set arrFLAILS(nC) = New clsFLAIL
        arrFLAILS(nC).bACTIVE = False
        nC = nC + 1
    Loop
    
    ' generate level monsters
    ' TODO
    'ReDim arrTOBEMONSTERS(0 To 5)
    'arrTOBEMONSTERS(0) = 1
    'arrTOBEMONSTERS(1) = 1
    'arrTOBEMONSTERS(2) = 1
    'arrTOBEMONSTERS(3) = 1
    'arrTOBEMONSTERS(4) = 1
    'arrTOBEMONSTERS(5) = 1
    
    generateMONSTERS (lCURRENTLEVEL * 20) + 10 ' generate monsters
    
    Dim currSTARTTIME As Currency
    Dim currCURRENTTIME As Currency
    Dim currFREQUENCY As Currency
    Dim dblTIMEBETWEENFRAMES As Double
    
    QueryPerformanceFrequency currFREQUENCY ' get the frequency of ticks
    dblTIMEBETWEENFRAMES = currFREQUENCY / FPS ' get time between frames needed to reach FPS
    Do While bEXIT = False And bFORCEEXIT = False
        QueryPerformanceCounter currCURRENTTIME ' get current time
        If currCURRENTTIME >= currSTARTTIME + dblTIMEBETWEENFRAMES Then ' if start time + time between frame = current time, then time for the next frame
            QueryPerformanceCounter currSTARTTIME ' store current time as new start time
            moveEVERYTHING ' move everything
            'drawEVERYTHING ' draw everything
            'checkWINLOOSE ' run win/loose code
        Else
            Sleep 3
        End If
        DoEvents ' do any events needed to be done
    Loop
    
    log "Stopping game..."
    
    If bFORCEEXIT = False Then ' if program not closing
        Dim lHIGHESTSCORE As Long
        
        ' allow shared win/loose code
        Dim sngMONEYMULTIPLYER As Single
        Dim strWINLOOSE As String
        
        If lCASTLECURRENTHEALTH <= 0 Then
            lCASTLECURRENTHEALTH = 1 ' reset current health, enough to continue if round winner doesn't buy health
            sngMONEYMULTIPLYER = 0.5 ' half money
            strWINLOOSE = "Loose" ' you lost
        Else
            sngMONEYMULTIPLYER = 1 ' all of money
            strWINLOOSE = "Win" ' you won
            lCURRENTLEVEL = lCURRENTLEVEL + 1 ' next level
        End If
        
        lWINNEROFROUND = -1
        ' add score and find winner
        nC = 0
        Do While nC < MAXCLIENTS
            If cCLIENTS(nC).connected = True Then ' if connected
                lMONEY = safeADDLONG(lMONEY, CLng(cCLIENTINFO(nC).lLEVELSCORE / sngMONEYMULTIPLYER)) ' add money
                If lHIGHESTSCORE < cCLIENTINFO(nC).lLEVELSCORE Or lWINNEROFROUND = -1 Then ' see if winner of round
                    lWINNEROFROUND = nC
                End If
            End If
            nC = nC + 1
        Loop
        
        broadcast "health", CStr(lCASTLECURRENTHEALTH)
        broadcast "moneyTotal", CStr(lMONEY)
        broadcast "nextLevel", CStr(lCURRENTLEVEL)
        
        nC = 0
        Do While nC < MAXCLIENTS
            If nC = lWINNEROFROUND Then
                cCLIENTS(nC).sendString "game", "stop" & strWINLOOSE & "Shop"
            Else
                cCLIENTS(nC).sendString "game", "stop" & strWINLOOSE
            End If
            nC = nC + 1
        Loop
        
        ' reset user vars
        nC = 0
        Do While nC < MAXCLIENTS
            If cCLIENTS(nC).connected = True Then
                cCLIENTINFO(nC).afterLevelReset
            End If
            nC = nC + 1
        Loop
    End If
    
    'frmSERVER.timerGAME.Enabled = True ' start game timer
End Sub

Sub broadcast(strCOMMAND As String, strTOSEND As String)
    Dim nC As Integer
    nC = 0
    Do While nC < MAXCLIENTS
        If cCLIENTS(nC).connected = True Then
            cCLIENTS(nC).sendString strCOMMAND, strTOSEND
            DoEvents
        End If
        nC = nC + 1
    Loop
End Sub

Function getMONSTERINFO(lMONSTERNUMBER As Long) As String
    getMONSTERINFO = CStr(lMONSTERNUMBER) & "~" & _
    CStr(arrMONSTERS(lMONSTERNUMBER).bACTIVE) & "~" & _
    CStr(arrMONSTERS(lMONSTERNUMBER).intTYPE) & "~" & _
    CStr(arrMONSTERS(lMONSTERNUMBER).sngX) & "~" & _
    CStr(arrMONSTERS(lMONSTERNUMBER).sngY) & "~" & _
    CStr(arrMONSTERS(lMONSTERNUMBER).sngMOVINGH) & "~" & _
    CStr(arrMONSTERS(lMONSTERNUMBER).intHEALTH)
End Function

Sub syncMONSTERS()
    Dim nC As Long
    nC = 0
    Dim strTOSEND As String
    Do While nC <= UBound(arrMONSTERS)
        If arrMONSTERS(nC).bACTIVE = True Then ' only sync if active
            If strTOSEND <> "" Then ' if not first
                strTOSEND = strTOSEND & "\" ' add separator
            End If
            strTOSEND = strTOSEND & getMONSTERINFO(nC)
        End If
        nC = nC + 1
    Loop
    broadcast "syncMon", strTOSEND
End Sub

Sub broadcastMONSTER(lMONSTERNUMBER As Long)
    Dim strTOSEND As String
    
    ' put monster data into string
    strTOSEND = getMONSTERINFO(lMONSTERNUMBER)
    
    ' send monster data
    broadcast "updateMon", strTOSEND
End Sub

Function getFLAILINFO(lFLAILNUMBER As Long, bCLEARGOTHROUGH As Boolean) As String
    strTOSEND = _
    CStr(lFLAILNUMBER) & "~" & _
    CStr(arrFLAILS(lFLAILNUMBER).bACTIVE) & "~" & _
    CStr(arrFLAILS(lFLAILNUMBER).sngX) & "~" & _
    CStr(arrFLAILS(lFLAILNUMBER).sngY) & "~" & _
    CStr(arrFLAILS(lFLAILNUMBER).sngMOVINGV) & "~" & _
    CStr(arrFLAILS(lFLAILNUMBER).sngMOVINGH) & "~" & _
    CStr(arrFLAILS(lFLAILNUMBER).intGOTHROUGH) & "~" & _
    CStr(bCLEARGOTHROUGH)
End Function

Sub syncFLAILS()
    Dim nC As Long
    nC = 0
    Dim strTOSEND As String
    Do While nC <= UBound(arrFLAILS)
        If arrFLAILS(nC).bACTIVE = True Then ' only sync if active
            If strTOSEND <> "" Then ' if not first
                strTOSEND = strTOSEND & "\" ' add separator
            End If
            strTOSEND = strTOSEND & getFLAILINFO(nC, False)
        End If
        nC = nC + 1
    Loop
    broadcast "syncFla", strTOSEND
End Sub

Sub broadcastFLAIL(lFLAILNUMBER As Long, bCLEARGOTHROUGH As Boolean)
    Dim strTOSEND As String
    strTOSEND = getFLAILINFO(lFLAILNUMBER, bCLEARGOTHROUGH)
    broadcast "updateFlail", strTOSEND
End Sub

Public Sub spawnMONSTER()
    Dim nC As Long
    If intCURRENTMONSTER <= UBound(arrTOBEMONSTERS) Then
        nC = 0
        Do While nC <= UBound(arrMONSTERS)
            If arrMONSTERS(nC).bACTIVE = False Then
                arrMONSTERS(nC).bACTIVE = True
                arrMONSTERS(nC).intTYPE = arrTOBEMONSTERS(intCURRENTMONSTER) 'Int(Rnd() * numberOfMonsters)
                
                arrMONSTERS(nC).sngX = Int(Rnd() * 2)
                If arrMONSTERS(nC).sngX = 0 Then
                    arrMONSTERS(nC).sngX = 0 - cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).lWIDTH
                    arrMONSTERS(nC).sngMOVINGH = 1 * cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).sngSPEED ' go left
                Else
                    arrMONSTERS(nC).sngX = windowX + cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).lWIDTH
                    arrMONSTERS(nC).sngMOVINGH = -1 * cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).sngSPEED ' go left
                End If
                arrMONSTERS(nC).sngY = cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).intSTARTINGY
                arrMONSTERS(nC).intHEALTH = cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).intMAXHEALTH
                
                Exit Do
            End If
            nC = nC + 1
        Loop
        intCURRENTMONSTER = intCURRENTMONSTER + 1
        broadcastMONSTER nC
    End If
End Sub

Sub loadONEMONSTERINFO(intNUMBER As Integer, imageNAME As String, lIMAGEWIDTH As Long, lIMAGEHEIGHT As Long, intPOINTCOST As Integer, intHEALTH As Integer, intATTACKPOWER As Integer, intSTARTINGY As Integer, sngSPEED As Single, intMONEYONHIT As Integer, intMONEYONKILL As Integer)
    cmontypeMONSTERINFO(intNUMBER).intPOINTCOST = intPOINTCOST
    cmontypeMONSTERINFO(intNUMBER).intMAXHEALTH = intHEALTH
    cmontypeMONSTERINFO(intNUMBER).intATTACKPOWER = intATTACKPOWER
    cmontypeMONSTERINFO(intNUMBER).sngSPEED = sngSPEED
    If intSTARTINGY = -1 Then ' default: ground
        cmontypeMONSTERINFO(intNUMBER).intSTARTINGY = landHEIGHT - lIMAGEHEIGHT
    Else
        cmontypeMONSTERINFO(intNUMBER).intSTARTINGY = intSTARTINGY
    End If
    cmontypeMONSTERINFO(intNUMBER).intMONEYADDEDHIT = intMONEYONHIT
    cmontypeMONSTERINFO(intNUMBER).intMONEYADDEDKILL = intMONEYONKILL
    cmontypeMONSTERINFO(intNUMBER).lWIDTH = lIMAGEWIDTH
    cmontypeMONSTERINFO(intNUMBER).lHEIGHT = lIMAGEHEIGHT
End Sub

Sub Main()
'    ' need to get info from a client
'    bHASINFO = False
    loadMONSTERINFO
    
    Dim nC As Integer
    nC = 0
    Do While nC < MAXCLIENTS
        cCLIENTS(nC).arrayID = nC
        nC = nC + 1
    Loop
    
    ReDim arrMONSTERS(0 To 99)
    nC = 0
    Do While nC <= UBound(arrMONSTERS)
        Set arrMONSTERS(nC) = New clsMONSTER
        nC = nC + 1
    Loop
    
    ReDim arrFLAILS(0 To 99)
    nC = 0
    Do While nC <= UBound(arrFLAILS)
        Set arrFLAILS(nC) = New clsFLAIL
        nC = nC + 1
    Loop
    
    frmSERVER.Show
End Sub
