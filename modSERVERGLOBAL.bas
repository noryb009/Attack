Attribute VB_Name = "modGLOBAL"
Global Const VERSION = "0.0.0.1a"
Global Const SERVER = True
Global Const onlineMODE = False
Global Const MAXCLIENTS = 4

'Global bHASINFO As Boolean

Global lPORT As Long

Global cCLIENTS(0 To MAXCLIENTS - 1) As New clsCONNECTION
Global cCLIENTINFO(0 To MAXCLIENTS - 1) As New clsCLIENTINFO
Global intPLAYERS As Integer

Global lCURRENTLEVEL As Long
Global lLEVELMONEY As Long

' level vars
Global arrTOBEMONSTERS() As Integer
Public intCURRENTMONSTER As Integer
Public intMONSTERSKILLED As Integer
Public intMONSTERSATTACKEDCASTLE As Integer
Public bEXIT As Boolean

Global arrMONSTERS() As clsMONSTER
Global arrFLAILS() As clsFLAIL

Public Sub log(strNEWLINE As String)
    frmSERVER.txtLOG.Text = strNEWLINE & vbCrLf & frmSERVER.txtLOG.Text
End Sub

Public Sub moveEVERYTHING()
    Dim lMOVESPEED As Long
    lMOVESPEED = 0.5 + ((lCURRENTLEVEL / 5) * intPLAYERS)
    
    Dim nC As Long
    
    ' spawn monsters
    Dim bSPAWN As Boolean
    bSPAWN = False
    If intCURRENTMONSTER = intMONSTERSKILLED + intMONSTERSATTACKEDCASTLE + (lCURRENTLEVEL \ 3) + intPLAYERS Then ' force if nobody on screen
        bSPAWN = True
    ElseIf Int(Rnd() * 200) < lCURRENTLEVEL * intPLAYERS And intCURRENTMONSTER <= UBound(arrTOBEMONSTERS) Then ' randomly if some monsters are waiting
        bSPAWN = True
    End If
    
    If bSPAWN = True Then
        spawnMONSTER
    End If
    
    ' move monsters
    nC = 0
    Do While nC <= UBound(arrMONSTERS)
        If arrMONSTERS(nC).bACTIVE = True Then
            arrMONSTERS(nC).sngX = arrMONSTERS(nC).sngX + lMOVESPEED * arrMONSTERS(nC).sngMOVINGH
            If (arrMONSTERS(nC).sngMOVINGH < 0 And arrMONSTERS(nC).sngX + cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).lWIDTH < 0) Or (arrMONSTERS(nC).sngMOVINGH > 0 And arrMONSTERS(nC).sngX > windowX) Then
                arrMONSTERS(nC).bACTIVE = False
                broadcastMONSTER nC
            ElseIf (arrMONSTERS(nC).sngMOVINGH > 0 And arrMONSTERS(nC).sngX + cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).lWIDTH > castleWALLLEFT) Or (arrMONSTERS(nC).sngMOVINGH < 0 And arrMONSTERS(nC).sngX < castleWALLRIGHT) Then 'attack
                lCASTLECURRENTHEALTH = lCASTLECURRENTHEALTH - cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).intATTACKPOWER
                intMONSTERSATTACKEDCASTLE = intMONSTERSATTACKEDCASTLE + 1
                arrMONSTERS(nC).bACTIVE = False
                If lCASTLECURRENTHEALTH <= 0 Then bEXIT = True
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
                    lLEVELMONEY = safeADDLONG(lLEVELMONEY, cmontypeMONSTERINFO(arrMONSTERS(intDELETEMONSTER).intTYPE).intMONEYADDEDKILL)
                Else
                    If intFLAILGOTHROUGH > 1 Then arrFLAILS(nC).addGOTHROUGH intDELETEMONSTER
                    lLEVELMONEY = safeADDLONG(lLEVELMONEY, cmontypeMONSTERINFO(arrMONSTERS(intDELETEMONSTER).intTYPE).intMONEYADDEDHIT)
                End If
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
    log "Starting game..."
    lLEVELMONEY = 0
    broadcast "game", "start" ' broadcast to clients that the game is starting
    
    ' reset level vars
    bEXIT = False
    lLEVELMONEY = 0
    intMONSTERSKILLED = 0
    intMONSTERSATTACKEDCASTLE = 0
    intCURRENTMONSTER = 0
    ReDim arrTOBEMONSTERS(0 To 0)
    
    ' generate level monsters
    ' TODO
    ReDim arrTOBEMONSTERS(0 To 5)
    arrTOBEMONSTERS(0) = 1
    arrTOBEMONSTERS(1) = 1
    arrTOBEMONSTERS(2) = 1
    arrTOBEMONSTERS(3) = 1
    arrTOBEMONSTERS(4) = 1
    arrTOBEMONSTERS(5) = 1
    
    frmSERVER.timerGAME.Enabled = True ' start game timer
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

Sub broadcastMONSTER(lMONSTERNUMBER As Long)
    Dim strTOSEND As String
    
    ' put monster data into string
    strTOSEND = CStr(lMONSTERNUMBER) & "~" & _
    CStr(arrMONSTERS(lMONSTERNUMBER).bACTIVE) & "~" & _
    CStr(arrMONSTERS(lMONSTERNUMBER).intTYPE) & "~" & _
    CStr(arrMONSTERS(lMONSTERNUMBER).sngX) & "~" & _
    CStr(arrMONSTERS(lMONSTERNUMBER).sngY) & "~" & _
    CStr(arrMONSTERS(lMONSTERNUMBER).sngMOVINGH) & "~" & _
    CStr(arrMONSTERS(lMONSTERNUMBER).intHEALTH)
    
    ' send monster data
    broadcast "updateMon", strTOSEND
End Sub

Sub broadcastFLAIL(lFLAILNUMBER As Long, bCLEARGOTHROUGH As Boolean)
    Dim strTOSEND As String
    
    If arrFLAILS(lFLAILNUMBER).bACTIVE = False Then
        log "Flail removed" ' TODO: remove
    End If
    
    strTOSEND = _
    CStr(lFLAILNUMBER) & "~" & _
    CStr(arrFLAILS(lFLAILNUMBER).bACTIVE) & "~" & _
    CStr(arrFLAILS(lFLAILNUMBER).sngX) & "~" & _
    CStr(arrFLAILS(lFLAILNUMBER).sngY) & "~" & _
    CStr(arrFLAILS(lFLAILNUMBER).sngMOVINGV) & "~" & _
    CStr(arrFLAILS(lFLAILNUMBER).sngMOVINGH) & "~" & _
    CStr(arrFLAILS(lFLAILNUMBER).intGOTHROUGH) & "~" & _
    CStr(bCLEARGOTHROUGH)
    broadcast "updateFlail", strTOSEND
End Sub

Public Sub spawnMONSTER()
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
    End If
End Sub

Sub loadONEMONSTERINFO(intNUMBER As Integer, imageNAME As String, lIMAGEWIDTH As Long, lIMAGEHEIGHT As Long, intHEALTH As Integer, intATTACKPOWER As Integer, intSTARTINGY As Integer, sngSPEED As Single, intMONEYONHIT As Integer, intMONEYONKILL As Integer)
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

Sub main()
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
    
    intPLAYERS = 0
    lCURRENTLEVEL = 0
    
    frmSERVER.Show
End Sub
