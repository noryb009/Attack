Attribute VB_Name = "modREQUESTHANDLER"
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

Sub sckDISCONNECTED(lARRAYID As Long, Optional bDISABLEQUIET As Boolean = False) ' handle client disconnection
    log cCLIENTS(lARRAYID).ip & " (" & cCLIENTINFO(lARRAYID).strNAME & ") disconnected." ' log that the client disconnected
    
    cCLIENTINFO(lARRAYID).reset ' reset client info for next client
    intPLAYERS = intPLAYERS - 1
    broadcastPLAYERLIST ' give clients updated player list
    
    If bDISABLEQUIET = False And cCLIENTINFO(lARRAYID).strNAME <> "" Then ' if had name
        broadcast "chat", formatCHATMSG("[") & formatCHATMSG(cCLIENTINFO(lARRAYID).strNAME, lPLAYERCOLOURS(lARRAYID)) & formatCHATMSG(" logged out]") ' alert users about logging out
    End If
    
    checkIFEVERYONEREADY ' see if everyone is ready
    
    If intPLAYERS = 0 Then ' nobody playing
        bFORCEEXIT = True ' stop server
        If lCASTLECURRENTHEALTH = 0 Then ' if out of health
             lCASTLECURRENTHEALTH = 1 ' set to 10 health
        End If
    End If
End Sub

Sub handleError(lARRAYID As Long, strDESCRIPTION As String) ' handle an error from clsCONNECTION
    log "Error from " & lARRAYID & ": " & strDESCRIPTION ' log error
    cCLIENTS(lARRAYID).disconnect
End Sub

Sub broadcastPLAYERLIST() ' turn the player list into a sendable string, then broadcast it
    Dim strTOSEND As String
    strTOSEND = "" ' clear string to send
    Dim nC As Integer
    nC = 0
    Do While nC < MAXCLIENTS ' for each client
        If nC <> 0 Then ' if not first name
            strTOSEND = strTOSEND & "~" ' add separator
        End If
        If cCLIENTS(nC).connected = True And cCLIENTINFO(nC).strNAME <> "" Then ' if client is connected
            strTOSEND = strTOSEND & CStr(cCLIENTINFO(nC).lLEVELSCORE) & "\" & CStr(cCLIENTINFO(nC).bREADY) & "\" & Replace(Replace(cCLIENTINFO(nC).strNAME, "&", "&amp;"), "~", "&tide;")  ' add player score, ready state, and escaped name to strTOSEND
        End If
        nC = nC + 1 ' next client
    Loop
    broadcast "playerList", strTOSEND ' send player list to clients
End Sub

Public Sub checkIFEVERYONEREADY()
    If intPLAYERS < 1 Then ' if nobody connected
        Exit Sub ' don't start game
    End If
    
    Dim nC As Integer
    nC = 0
    
    If bPLAYING = False Then ' if game isn't already started
        ' see if everyone is ready
        Do While nC < MAXCLIENTS ' for each client
            If cCLIENTS(nC).connected = True Then ' if client is connected
                If cCLIENTINFO(nC).bREADY = False Then ' if client isn't ready
                    Exit Sub ' someone not ready, exit sub
                End If
            End If
            nC = nC + 1 ' next client
        Loop
        startGAME ' everyone is ready, start
    Else ' game is currently being played
        Dim bSHOULDSYNC As Boolean
        bSHOULDSYNC = False ' shouldn't sync yet
        Do While nC < MAXCLIENTS ' for each client
            If cCLIENTS(nC).connected = True Then ' if client is connected
                If cCLIENTINFO(nC).bREADY = True Then ' if client is ready
                    cCLIENTS(nC).sendString "game", "start" ' send into game
                    bSHOULDSYNC = True ' client added, should sync
                End If
            End If
            nC = nC + 1 ' next client
        Loop
        If bSHOULDSYNC = True Then ' if client was added, should sync
            syncMONSTERS ' sync monsters with everybody
            syncFLAILS ' sync flails with everybody
        End If
        Exit Sub
    End If
End Sub

Sub spawnFLAIL(lARRAYID As Long, strSTATS As String)
    Dim arrstrSTATS() As String
    arrstrSTATS = Split(strSTATS, "~") ' split stats
    
    Dim lFLAILSPOT As Long
    lFLAILSPOT = -1 ' flail spot in array to put data
    Dim nC As Integer
    Do While nC < lFLAILARRAYSIZE ' for each flail spot in array
        If arrFLAILS(nC).bACTIVE = False Then ' if not active
            lFLAILSPOT = nC ' remember number
            Exit Do ' found empty spot, continue
        End If
        nC = nC + 1 ' next flail spot
    Loop
    
    If lFLAILSPOT = -1 Then ' no room left in flail array
        Exit Sub ' exit
    End If
    
    arrFLAILS(lFLAILSPOT).lOWNER = lARRAYID ' store owner
    arrFLAILS(lFLAILSPOT).bACTIVE = CBool(arrstrSTATS(0)) ' store if active
    arrFLAILS(lFLAILSPOT).sngX = CSng(arrstrSTATS(1)) ' store X location
    arrFLAILS(lFLAILSPOT).sngY = CSng(arrstrSTATS(2)) ' store Y location
    arrFLAILS(lFLAILSPOT).sngMOVINGV = CSng(arrstrSTATS(3)) ' store vertical speed
    arrFLAILS(lFLAILSPOT).sngMOVINGH = CSng(arrstrSTATS(4)) ' store horizontal speed
    arrFLAILS(lFLAILSPOT).intGOTHROUGH = CInt(arrstrSTATS(5)) ' store go through left
    If CBool(arrstrSTATS(6)) = True Then ' if we should clear go through
        arrFLAILS(lFLAILSPOT).clearWENTTHROUGH ' clear go through
    End If
    
    broadcastFLAIL lFLAILSPOT, True ' broadcast the new flail
End Sub

Public Sub handleREQUEST(lARRAYID As Long, strCOMMAND As String, strDESCRIPTION As String) ' handle a request from a client
    Dim nC As Integer
    Select Case strCOMMAND ' which command is it
        Case "VERSION" ' check version
            If strDESCRIPTION = VERSION Then ' if same version
                log "Correct version received from " & cCLIENTS(lARRAYID).ip ' log that correct version was received
                cCLIENTS(lARRAYID).sendString "login" ' ask for user to log in
            Else
                log "Incorrect version received from " & cCLIENTS(lARRAYID).ip ' log that incorrent version was received
                cCLIENTS(lARRAYID).sendString "DISCONNECT", "version mismatch" ' send error
                cCLIENTS(lARRAYID).connected = False ' not connected anymore
            End If
        Case "login" ' user is logging in
            If strDESCRIPTION = "" Or Len(strDESCRIPTION) > 15 Then ' invalid username
                log cCLIENTS(lARRAYID).ip & " tried to log in as " & strNAME ' log invalid login
                cCLIENTS(lARRAYID).sendString "DISCONNECT", "invalid name" ' alert user that the name is invalid
                cCLIENTS(lARRAYID).connected = False ' not connected anymore
                Exit Sub ' exit
            End If
            
            ' see if username already used
            nC = 0
            Do While nC < MAXCLIENTS ' for each client
                If cCLIENTS(nC).connected = True And nC <> lARRAYID Then ' if user if connected and not the current client
                    If UCase(cCLIENTINFO(nC).strNAME) = UCase(strDESCRIPTION) Then ' if the name is the same
                        cCLIENTS(lARRAYID).sendString "DISCONNECT", "name already in use" ' disconnect user
                        cCLIENTS(lARRAYID).connected = False ' not connected anymore
                        Exit Sub ' exit
                    End If
                End If
                nC = nC + 1 ' next client
            Loop
            cCLIENTINFO(lARRAYID).strNAME = strDESCRIPTION ' store name
            log cCLIENTS(lARRAYID).ip & " logged in as " & cCLIENTINFO(lARRAYID).strNAME ' log what the user logged in as
            cCLIENTS(lARRAYID).sendString "login", "success" ' tell client they are logged in
            DoEvents ' after sendstring, run doevents to send to client
            ' send user current stats
            cCLIENTS(lARRAYID).sendString "flaPower", CStr(intFLAILPOWER) ' sync flail power
            DoEvents ' after sendstring, run doevents to send to client
            cCLIENTS(lARRAYID).sendString "flaGoThrough", CStr(intFLAILGOTHROUGH) ' sync flail gothrough
            DoEvents ' after sendstring, run doevents to send to client
            cCLIENTS(lARRAYID).sendString "flaAmount", CStr(intFLAILAMOUNT) ' sync flail amount
            DoEvents ' after sendstring, run doevents to send to client
            cCLIENTS(lARRAYID).sendString "moneyTotal", CStr(lMONEY) ' sync money
            DoEvents ' after sendstring, run doevents to send to client
            cCLIENTS(lARRAYID).sendString "health", CStr(lCASTLECURRENTHEALTH) ' sync health
            DoEvents ' after sendstring, run doevents to send to client
            cCLIENTS(lARRAYID).sendString "maxHealth", CStr(lCASTLEMAXHEALTH) ' sync max health
            DoEvents ' after sendstring, run doevents to send to client
            cCLIENTS(lARRAYID).sendString "monstersLeft", getMONSTERSLEFT ' sync number of monsters left
            DoEvents ' after sendstring, run doevents to send to client
            cCLIENTS(lARRAYID).sendString "nextLevel", CStr(lCURRENTLEVEL) ' sync current level
            DoEvents ' after sendstring, run doevents to send to client
            
            ' give clients updated player list
            broadcastPLAYERLIST ' sync player list with all clients
            DoEvents ' after broadcast, run doevents to send to client
            ' broadcast new user
            broadcast "chat", formatCHATMSG("[") & formatCHATMSG(cCLIENTINFO(lARRAYID).strNAME, lPLAYERCOLOURS(lARRAYID)) & formatCHATMSG(" logged in]") ' broadcast that a user logged in
        Case "newFla" ' user created a flail
            If strDESCRIPTION = "" Then ' if empty
                log "Empty newFla received from " & cCLIENTS(lARRAYID).ip ' log the bad command
            Else
                spawnFLAIL lARRAYID, strDESCRIPTION ' spawn the flail
            End If
        Case "chat" ' user is talking
            log "Chat: " & cCLIENTINFO(lARRAYID).strNAME & ": " & strDESCRIPTION ' log who said what
            broadcast "chat", formatCHATMSG(cCLIENTINFO(lARRAYID).strNAME, lPLAYERCOLOURS(lARRAYID)) & formatCHATMSG(": " & strDESCRIPTION) ' send message to other clients
        Case "ready" ' user is/isn't ready
            cCLIENTINFO(lARRAYID).bREADY = CBool(strDESCRIPTION) ' set as ready/not ready
            broadcast "readyState", CStr(lARRAYID) & "\" & CStr(cCLIENTINFO(lARRAYID).bREADY) ' update clients' ready state
            If CBool(strDESCRIPTION) = True Then ' if ready
                broadcast "chat", formatCHATMSG("[") & formatCHATMSG(cCLIENTINFO(lARRAYID).strNAME, lPLAYERCOLOURS(lARRAYID)) & formatCHATMSG(" is ready]") ' broadcast to clients that player is ready
                checkIFEVERYONEREADY ' check if everyone is ready to start game
            Else ' user isn't ready anymore
                broadcast "chat", formatCHATMSG("[") & formatCHATMSG(cCLIENTINFO(lARRAYID).strNAME, lPLAYERCOLOURS(lARRAYID)) & formatCHATMSG(" is no longer ready]") ' broadcast to clients that player isn't ready
            End If
        Case "heal" ' user is buying heal
            Dim strHEALPARTS() As String
            strHEALPARTS = Split(strDESCRIPTION, "~", 2) ' split into (cost, amount healed)
            If UBound(strHEALPARTS) = 1 Then ' if enough parts
                If lMONEY - CLng(strHEALPARTS(0)) >= 0 Then ' if you have enough money
                    lMONEY = lMONEY - CLng(strHEALPARTS(0)) ' take away money
                    lCASTLECURRENTHEALTH = lCASTLECURRENTHEALTH + CLng(strHEALPARTS(1)) ' heal
                End If
                broadcast "moneyTotal", CStr(lMONEY) ' broadcast new money
                broadcast "health", CLng(lCASTLECURRENTHEALTH) ' update other clients with new health
                broadcast "chat", formatCHATMSG("[") & formatCHATMSG(cCLIENTINFO(lARRAYID).strNAME, lPLAYERCOLOURS(lARRAYID)) & formatCHATMSG(" bought more health. You now have " & addZEROIFNOTZERO(lCASTLECURRENTHEALTH) & "/" & lCASTLEMAXHEALTH & "0 health.]")
            Else ' bad command
                log "Bad 'heal' command from " & cCLIENTS(lARRAYID).ip & ": " & strDESCRIPTION ' log the bad command
            End If
        Case "addHealth" ' user is buying more max health
            Dim strADDHEALTHPARTS() As String
            strADDHEALTHPARTS = Split(strDESCRIPTION, "~", 2) ' split into (cost, amount to raise max health)
            If UBound(strADDHEALTHPARTS) = 1 Then ' if enough parts
                If lMONEY - CLng(strADDHEALTHPARTS(0)) >= 0 Then ' if you have enough money
                    lMONEY = lMONEY - CLng(strADDHEALTHPARTS(0)) ' cost
                    lCASTLEMAXHEALTH = lCASTLEMAXHEALTH + CLng(strADDHEALTHPARTS(1)) ' more health
                    lCASTLECURRENTHEALTH = lCASTLECURRENTHEALTH + CLng(strADDHEALTHPARTS(1)) ' heal
                End If
                broadcast "moneyTotal", CStr(lMONEY) ' broadcast new money
                broadcast "maxHealth", CLng(lCASTLEMAXHEALTH) ' broadcast new max health
                broadcast "health", CLng(lCASTLECURRENTHEALTH) ' broadcast new health
                broadcast "chat", formatCHATMSG("[") & formatCHATMSG(cCLIENTINFO(lARRAYID).strNAME, lPLAYERCOLOURS(lARRAYID)) & formatCHATMSG(" bought more max health. You now have " & lCASTLEMAXHEALTH & "0 health.]")
            Else ' bad command
                log "Bad 'addHealth' command from " & cCLIENTS(lARRAYID).ip & ": " & strDESCRIPTION ' log bad command
            End If
        Case "buy" ' user is buying a flail upgrade
            Dim strBUYPARTS() As String
            strBUYPARTS = Split(strDESCRIPTION, "~", 2) ' split into (upgrade what, cost)
            If UBound(strBUYPARTS) = 1 Then ' if enough parts
                If lMONEY - CLng(strBUYPARTS(1)) >= 0 Then ' if you have enough money
                    If strBUYPARTS(0) = "power" Then ' if buying power
                        intFLAILPOWER = intFLAILPOWER + 1 ' increase power
                        broadcast "flaPower", CStr(intFLAILPOWER) ' broadcast new power
                        broadcast "chat", formatCHATMSG("[") & formatCHATMSG(cCLIENTINFO(lARRAYID).strNAME, lPLAYERCOLOURS(lARRAYID)) & formatCHATMSG(" bought more flail power. You now have " & intFLAILPOWER & " flail power.]")
                    ElseIf strBUYPARTS(0) = "goThrough" Then ' if buying go through
                        intFLAILGOTHROUGH = intFLAILGOTHROUGH + 1 ' increase go through
                        broadcast "flaGoThrough", CStr(intFLAILGOTHROUGH) ' broadcast go through
                        broadcast "chat", formatCHATMSG("[") & formatCHATMSG(cCLIENTINFO(lARRAYID).strNAME, lPLAYERCOLOURS(lARRAYID)) & formatCHATMSG(" bought more flail piercing power. You now have " & intFLAILGOTHROUGH & " flail piercing power.]")
                    Else 'If strBUYPARTS(0) = "amount" Then ' if buying amount
                        intFLAILAMOUNT = intFLAILAMOUNT + 1 ' increase amount
                        broadcast "flaAmount", CStr(intFLAILAMOUNT) ' broadcast amount
                        broadcast "chat", formatCHATMSG("[") & formatCHATMSG(cCLIENTINFO(lARRAYID).strNAME, lPLAYERCOLOURS(lARRAYID)) & formatCHATMSG(" bought more another flail. You now have " & intFLAILAMOUNT & " flails.]")
                    End If
                    lMONEY = lMONEY - CLng(strBUYPARTS(1)) ' money spent
                Else
                    broadcast "flaPower", CStr(intFLAILPOWER) ' broadcast new power
                    broadcast "flaGoThrough", CStr(intFLAILGOTHROUGH) ' broadcast go through
                    broadcast "flaAmount", CStr(intFLAILAMOUNT) ' broadcast amount
                End If
                broadcast "moneyTotal", CStr(lMONEY) ' broadcast new money
            Else ' bad command
                log "Bad 'buy' command from " & cCLIENTS(lARRAYID).ip & ": " & strDESCRIPTION ' log bad command
            End If
        Case Else ' not one of the above commands
            log "Unknown command from " & cCLIENTS(lARRAYID).ip & ": " & strCOMMAND ' log unknown command
    End Select
End Sub
