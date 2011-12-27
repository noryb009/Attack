Attribute VB_Name = "modREQUESTHANDLER"
Sub sckDISCONNECTED(lARRAYID As Long)
    log cCLIENTS(lARRAYID).ip & " (" & cCLIENTINFO(lARRAYID).strNAME & ") disconnected."
    
    If cCLIENTINFO(lARRAYID).strNAME <> "" Then
        broadcast "chat", "[" & cCLIENTINFO(lARRAYID).strNAME & " logged out]"
    End If
    
    cCLIENTINFO(lARRAYID).reset
    intPLAYERS = intPLAYERS - 1
    broadcastPLAYERLIST ' give clients updated player list
    checkIFEVERYONEREADY ' see if everyone is ready
    
    If intPLAYERS = 0 Then ' nobody playing
        bFORCEEXIT = True
    End If
End Sub

Sub handleError(lARRAYID As Long, strDESCRIPTION As String)
    log "Error from " & lARRAYID & ":" & strDESCRIPTION
End Sub

Sub broadcastPLAYERLIST() ' turn the player list into a sendable string, then broadcast it
    Dim strTOSEND As String
    strTOSEND = ""
    Dim nC As Integer
    nC = 0
    Do While nC < MAXCLIENTS
        If cCLIENTS(nC).connected = True And cCLIENTINFO(nC).strNAME <> "" Then
            If strTOSEND <> "" Then ' if not first name
                strTOSEND = strTOSEND & "~" ' add separator
            End If
            strTOSEND = strTOSEND & Replace(Replace(cCLIENTINFO(nC).strNAME, "&", "&amp;"), "~", "&tide;") ' "&" to "&amp;", "~" to "&tide;", add to strTOSEND
        End If
        nC = nC + 1
    Loop
    broadcast "playerList", strTOSEND ' send string
End Sub

Public Sub checkIFEVERYONEREADY()
    If intPLAYERS = 0 Then ' if nobody connected
        Exit Sub ' don't start game
    End If
    
    ' see if everyone is ready
    nC = 0
    Do While nC < MAXCLIENTS
        If cCLIENTS(nC).connected = True Then
            If cCLIENTINFO(nC).bREADY = False Then
                Exit Sub ' someone not ready, exit sub
            End If
        End If
        nC = nC + 1
    Loop
    
    startGAME ' everyone is ready, start
End Sub

Sub spawnFLAIL(lARRAYID As Long, strSTATS As String)
    Dim arrstrSTATS() As String
    arrstrSTATS = Split(strSTATS, "~") ' split stats
    
    Dim lFLAILSPOT As Long
    lFLAILSPOT = -1 ' flail spot in array to put data
    Dim nC As Integer
    Do While nC <= UBound(arrFLAILS) ' for each flail spot in array
        If arrFLAILS(nC).bACTIVE = False Then ' if not active
            lFLAILSPOT = nC ' remember number
            Exit Do
        End If
        nC = nC + 1
    Loop
    
    If lFLAILSPOT = -1 Then ' no room left in flail array
        ReDim Preserve arrFLAILS(0 To UBound(arrFLAILS) + 1) ' make array 1 bigger
        Set arrFLAILS(UBound(arrFLAILS)) = New clsFLAIL
        lFLAILSPOT = UBound(arrFLAILS)
    End If
    
    arrFLAILS(lFLAILSPOT).lOWNER = lARRAYID
    arrFLAILS(lFLAILSPOT).bACTIVE = CBool(arrstrSTATS(0))
    arrFLAILS(lFLAILSPOT).sngX = CSng(arrstrSTATS(1))
    arrFLAILS(lFLAILSPOT).sngY = CSng(arrstrSTATS(2))
    arrFLAILS(lFLAILSPOT).sngMOVINGV = CSng(arrstrSTATS(3))
    arrFLAILS(lFLAILSPOT).sngMOVINGH = CSng(arrstrSTATS(4))
    arrFLAILS(lFLAILSPOT).intGOTHROUGH = CInt(arrstrSTATS(5))
    If CBool(arrstrSTATS(6)) = True Then
        arrFLAILS(lFLAILSPOT).clearWENTTHROUGH
    End If
    
    broadcastFLAIL lFLAILSPOT, True
End Sub

'Sub getMONINFO(lARRAYID As Long, strINFO As String)
'    If bHASINFO = False Then ' don't parse again if got from another client
'        Dim arrstrINFO() As String
'        arrstrINFO = Split(strINFO, "\")
'        If UBound(arrstrINFO) + 1 <> numberOfMonsters Then
'                log "Error loading monster info."
'                cCLIENTS(lARRAYID).disconnect
'        End If
'        'TODO: check
'        Dim oneMONSTERINFO() As String
'        Dim nC As Integer
'        nC = 0
'        Do While nC < numberOfMonsters
'            oneMONSTERINFO = Split(arrstrINFO(nC), "~")
'            If UBound(oneMONSTERINFO) <> 5 Then
'                log "Error loading monster info."
'                cCLIENTS(lARRAYID).disconnect
'            End If
'            cmontypeMONSTERINFO(nC).intATTACKPOWER = CInt(oneMONSTERINFO(0))
'            cmontypeMONSTERINFO(nC).intMAXHEALTH = CInt(oneMONSTERINFO(1))
'            cmontypeMONSTERINFO(nC).intMONEYADDEDHIT = CInt(oneMONSTERINFO(2))
'            cmontypeMONSTERINFO(nC).intMONEYADDEDKILL = CInt(oneMONSTERINFO(3))
'            cmontypeMONSTERINFO(nC).intSTARTINGY = CInt(oneMONSTERINFO(4))
'            cmontypeMONSTERINFO(nC).sngSPEED = CSng(oneMONSTERINFO(5))
'            nC = nC + 1
'        Loop
'    End If
'End Sub

Public Sub handleREQUEST(lARRAYID As Long, strCOMMAND As String, strDESCRIPTION As String)
    Dim nC As Integer
    Select Case strCOMMAND
        Case "VERSION"
            If strDESCRIPTION = VERSION Then
                log "Correct version received from " & cCLIENTS(lARRAYID).ip
                cCLIENTS(lARRAYID).sendString "login"
'                If bHASINFO = False Then
'                    cCLIENTS(lARRAYID).sendString "monInfo"
'                End If
            Else
                log "Incorrect version received from " & cCLIENTS(lARRAYID).ip
                cCLIENTS(lARRAYID).sendString "DISCONNECT", "version mismatch"
                cCLIENTS(lARRAYID).connected = False
            End If
        Case "login"
            If strDESCRIPTION = "" Or Len(strDESCRIPTION) > 25 Then
                log cCLIENTS(lARRAYID).ip & " tried to log in as " & strNAME
                cCLIENTS(lARRAYID).sendString "DISCONNECT", "invalid name"
                cCLIENTS(lARRAYID).connected = False
                Exit Sub
            End If
            
            nC = 0
            Do While nC < MAXCLIENTS
                If cCLIENTS(nC).connected = True And nC <> lARRAYID Then
                    If cCLIENTINFO(nC).strNAME = strDESCRIPTION Then
                        cCLIENTS(lARRAYID).sendString "DISCONNECT", "name already in use"
                        cCLIENTS(lARRAYID).connected = False
                        Exit Sub
                    End If
                End If
                nC = nC + 1
            Loop
            cCLIENTINFO(lARRAYID).strNAME = strDESCRIPTION
            log cCLIENTS(lARRAYID).ip & " logged in as " & cCLIENTINFO(lARRAYID).strNAME
            cCLIENTS(lARRAYID).sendString "login", "success" ' tell client they are logged in
            DoEvents
            ' send user current stats
            cCLIENTS(lARRAYID).sendString "flaPower", CStr(intFLAILPOWER)
            DoEvents
            cCLIENTS(lARRAYID).sendString "flaGoThrough", CStr(intFLAILGOTHROUGH)
            DoEvents
            cCLIENTS(lARRAYID).sendString "flaAmount", CStr(intFLAILAMOUNT)
            DoEvents
            cCLIENTS(lARRAYID).sendString "moneyTotal", CStr(lMONEY)
            DoEvents
            cCLIENTS(lARRAYID).sendString "health", CStr(lCASTLECURRENTHEALTH)
            DoEvents
            cCLIENTS(lARRAYID).sendString "maxHealth", CStr(lCASTLEMAXHEALTH)
            DoEvents
            cCLIENTS(lARRAYID).sendString "nextLevel", CStr(lCURRENTLEVEL)
            DoEvents
            
            ' give clients updated player list
            broadcastPLAYERLIST
            DoEvents
            ' broadcast new user
            broadcast "chat", "[" & cCLIENTINFO(lARRAYID).strNAME & " logged in]"
        Case "newFla"
            If strDESCRIPTION = "" Then
                log "Empty newFla received from " & cCLIENTS(lARRAYID).ip
            Else
                log "newFla received from " & cCLIENTS(lARRAYID).ip
                spawnFLAIL lARRAYID, strDESCRIPTION
            End If
'        Case "monInfo"
'            If strDESCRIPTION = "" Then
'                log "Empty monInfo received from " & cCLIENTS(lARRAYID).ip
'            Else
'                log "monInfo received from " & cCLIENTS(lARRAYID).ip
'                getMONINFO lARRAYID, strDESCRIPTION
'            End If
        Case "chat"
            log "Chat: " & cCLIENTINFO(lARRAYID).strNAME & ": " & strDESCRIPTION
            broadcast "chat", cCLIENTINFO(lARRAYID).strNAME & ": " & strDESCRIPTION
        Case "ready"
            If CBool(strDESCRIPTION) = True Then
                cCLIENTINFO(lARRAYID).bREADY = True
                broadcast "chat", "[" & cCLIENTINFO(lARRAYID).strNAME & " is ready]"
                checkIFEVERYONEREADY
            Else
                broadcast "chat", "[" & cCLIENTINFO(lARRAYID).strNAME & " is no longer ready]"
                cCLIENTINFO(lARRAYID).bREADY = False
            End If
        Case "heal"
            Dim strHEALPARTS() As String
            strHEALPARTS = Split(strDESCRIPTION, "~", 2)
            If UBound(strHEALPARTS) = 1 Then
                lMONEY = lMONEY - CLng(strHEALPARTS(0)) ' cost
                broadcast "moneyTotal", CStr(lMONEY) ' broadcast new money
                lCASTLECURRENTHEALTH = lCASTLECURRENTHEALTH + CLng(strHEALPARTS(1)) ' heal
                broadcast "health", CLng(lCASTLECURRENTHEALTH)
            Else ' bad command
                log "Bad 'heal' command from " & cCLIENTS(lARRAYID).ip & ": " & strDESCRIPTION
            End If
        Case "addHealth"
            Dim strADDHEALTHPARTS() As String
            strADDHEALTHPARTS = Split(strDESCRIPTION, "~", 2)
            If UBound(strADDHEALTHPARTS) = 1 Then
                lMONEY = lMONEY - CLng(strADDHEALTHPARTS(0)) ' cost
                broadcast "moneyTotal", CStr(lMONEY) ' broadcast new money
                lCASTLEMAXHEALTH = lCASTLEMAXHEALTH + CLng(strADDHEALTHPARTS(1)) ' more health
                broadcast "maxHealth", CLng(lCASTLEMAXHEALTH)
                lCASTLECURRENTHEALTH = lCASTLECURRENTHEALTH + CLng(strADDHEALTHPARTS(1)) ' heal
                broadcast "health", CLng(lCASTLECURRENTHEALTH)
            Else ' bad command
                log "Bad 'addHealth' command from " & cCLIENTS(lARRAYID).ip & ": " & strDESCRIPTION
            End If
        Case "buy"
            Dim strBUYPARTS() As String
            strBUYPARTS = Split(strDESCRIPTION, "~", 2)
            If UBound(strBUYPARTS) = 1 Then
                If strBUYPARTS(0) = "power" Then
                    intFLAILPOWER = intFLAILPOWER + 1
                    broadcast "flaPower", CStr(intFLAILPOWER)
                ElseIf strBUYPARTS(0) = "goThrough" Then
                    intFLAILGOTHROUGH = intFLAILGOTHROUGH + 1
                    broadcast "flaGoThrough", CStr(intFLAILGOTHROUGH)
                Else 'If strBUYPARTS(0) = "amount" Then
                    intFLAILAMOUNT = intFLAILAMOUNT + 1
                    broadcast "flaAmount", CStr(intFLAILAMOUNT)
                End If
                lMONEY = lMONEY - CLng(strBUYPARTS(1)) ' money spent
                broadcast "moneyTotal", CStr(lMONEY) ' broadcast new money
            Else ' bad command
                log "Bad 'buy' command from " & cCLIENTS(lARRAYID).ip & ": " & strDESCRIPTION
            End If
        Case Else
            log "Unknown command from " & cCLIENTS(lARRAYID).ip & ": " & strCOMMAND
    End Select
End Sub
