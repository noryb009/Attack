Attribute VB_Name = "modREQUESTHANDLER"
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

Sub sckDISCONNECTED(lARRAYID As Long, Optional bMESSAGE As Boolean = True) ' somebody disconnected
    Select Case currentSTATE ' what form is open
        Case "lobby" ' if lobby open
            Unload frmLOBBY ' unload lobby
        Case "lobbyShop" ' if lobby and shop open
            Unload frmLOBBY ' unload lobby
            Unload frmSTORE ' unload shop
        Case "playing" ' if playing
            Unload frmATTACK ' unload playing form
    End Select
    If onlineMODE = True Then ' if not already messaged
        If bMESSAGE = True Then MsgBox "Disconnected from host!", vbOKOnly, programNAME ' message that you disconnected
        onlineMODE = False ' not online anymore
    End If
    If currentSTATE <> "" Then ' if not already in new game form
        frmNEWGAME.Show ' show new game form
    End If
End Sub

Sub handleError(lARRAYID As Long, strDESCRIPTION As String) ' handle error from clsCONNECTION
    cSERVER(0).disconnect ' disconnect from server
End Sub

Sub updateMONSTER(strSTATS As String) ' monster update from server
    Dim arrstrSTATS() As String
    arrstrSTATS = Split(strSTATS, "~") ' get different data parts
    
    Dim lSPOT As Long
    lSPOT = CLng(arrstrSTATS(0)) ' get monster spot in arrMONSTERS
    
    Do While lSPOT >= lMONSTERARRAYSIZE ' if bigger then array size
        Exit Sub ' exit
    Loop
    
    ' copy new values
    arrMONSTERS(lSPOT).bACTIVE = CBool(arrstrSTATS(1))
    arrMONSTERS(lSPOT).currentFRAME = 0
    arrMONSTERS(lSPOT).intTYPE = CInt(arrstrSTATS(2))
    arrMONSTERS(lSPOT).sngX = CSng(arrstrSTATS(3))
    arrMONSTERS(lSPOT).sngY = CSng(arrstrSTATS(4))
    arrMONSTERS(lSPOT).sngMOVINGH = CSng(arrstrSTATS(5))
    arrMONSTERS(lSPOT).sngMOVINGV = cmontypeMONSTERINFO(arrMONSTERS(lSPOT).intTYPE).sngYSPEED
    arrMONSTERS(lSPOT).intHEALTH = CLng(arrstrSTATS(6))
End Sub

Sub syncMONSTERS(strALLMONINFO As String) ' sync all the monsters
    Dim nC As Integer
    nC = 0
    ' deactivate all monsters
    Do While nC < lMONSTERARRAYSIZE ' for each monster
        arrMONSTERS(nC).bACTIVE = False ' not active
        nC = nC + 1 ' next monster
    Loop
    
    Dim strONEMONINFO() As String
    strONEMONINFO = Split(strALLMONINFO, "\") ' split into active monsters
    If strONEMONINFO(0) <> "" Then ' if array isn't empty
        nC = 0
        Do While nC <= UBound(strALLMONS) ' for each new monster info
            updateMONSTER strONEMONINFO(nC) ' add monster
            nC = nC + 1 ' next new monster
        Loop
    End If
End Sub

Sub updateFLAIL(strSTATS As String) ' flail update from server
    Dim arrstrSTATS() As String
    arrstrSTATS = Split(strSTATS, "~") ' get different data parts
    
    If UBound(arrstrSTATS) <> 8 Then ' bad command
        Exit Sub ' exit
    End If
    
    Dim lSPOT As Long
    lSPOT = CLng(arrstrSTATS(0)) ' get flail spot in arrFLAILS
    
    Do While lSPOT >= lFLAILARRAYSIZE ' if bigger then current array size
        Exit Sub ' exit
    Loop
    
    ' copy new values
    arrFLAILS(lSPOT).bACTIVE = CBool(arrstrSTATS(1))
    arrFLAILS(lSPOT).lOWNER = CInt(arrstrSTATS(2))
    arrFLAILS(lSPOT).sngX = CSng(arrstrSTATS(3))
    arrFLAILS(lSPOT).sngY = CSng(arrstrSTATS(4))
    arrFLAILS(lSPOT).sngMOVINGV = CSng(arrstrSTATS(5))
    arrFLAILS(lSPOT).sngMOVINGH = CSng(arrstrSTATS(6))
    arrFLAILS(lSPOT).intGOTHROUGH = CInt(arrstrSTATS(7))
    If CBool(arrstrSTATS(8)) = True Then ' if we should clear go through
        arrFLAILS(lSPOT).clearWENTTHROUGH ' clear go through
    End If
End Sub

Sub syncFLAILS(strALLFLAINFO As String) ' sync all the flails from the server
    Dim nC As Integer
    nC = 0
    ' deactivate all flails
    Do While nC < lFLAILARRAYSIZE ' for each flail
        arrFLAILS(nC).bACTIVE = False ' deactivate flail
        nC = nC + 1 ' next flail
    Loop
    
    Dim strONEFLAINFO() As String
    strONEFLAINFO = Split(strALLFLAINFO, "\") ' split into all flails
    If strONEFLAINFO(0) <> "" Then ' if array isn't empty
        nC = 0
        Do While nC <= UBound(strONEFLAINFO) ' for each new flail info
            updateFLAIL strONEFLAINFO(nC) ' add flail
            nC = nC + 1 ' next new flail
        Loop
    End If
End Sub

Public Sub handleCHAT(strMESSAGE As String)
    If currentSTATE <> "lobby" And currentSTATE <> "lobbyShop" And currentSTATE <> "playing" Then ' if not in lobby or game area
        Exit Sub ' exit
    End If
    
    Dim strMESSAGEPARTS() As String ' message, separated by colour switches ("colour\message")
    Dim strPARTSOFMESSAGEPART() As String ' strMESSAGEPARTS parts ("colour", "message")
    Dim lCOLOUR As Long ' colour part of message
    Dim bBOLD As Boolean ' bold message or not
    Dim strONEMESSAGEPART As String ' text part of message
    
    Dim nC As Integer
    
    ' before loop
    If currentSTATE = "playing" Then ' if playing
        ' bump old messages in chat log
        nC = 0
        Do While nC < UBound(strCHATLOG) ' for each (not last) chat log place
            strCHATLOG(nC) = strCHATLOG(nC + 1) ' move message below up
            nC = nC + 1 ' next chat log spot
        Loop
        strCHATLOG(UBound(strCHATLOG)) = "" ' clear current chat log place
    Else ' in lobby
        If frmLOBBY.rtbCHATLOG.Text <> "" Then ' if not empty
            frmLOBBY.rtbCHATLOG.SelStart = Len(frmLOBBY.rtbCHATLOG.Text)
            frmLOBBY.rtbCHATLOG.SelText = vbCrLf ' add newline
        End If
    End If
    
    strMESSAGEPARTS = Split(strMESSAGE, "~") ' split message into colour and text parts
    
    nC = 0
    Do While nC <= UBound(strMESSAGEPARTS) ' for each message part
        If nC = UBound(strMESSAGEPARTS) And strMESSAGEPARTS(nC) = "" Then ' if nothing on last message part
            Exit Do ' exit
        End If
        If InStr(strMESSAGEPARTS(nC), "\") <> 0 Then ' if includes colour info
            strPARTSOFMESSAGEPART = Split(strMESSAGEPARTS(nC), "\", 3) ' strPARTSOFMESSAGEPART="colour", "message"
            If UBound(strPARTSOFMESSAGEPART) <> 2 Then ' if not enough parts
                lCOLOUR = vbBlack ' default: black
                bBOLD = False ' default: not bold
                strONEMESSAGEPART = strMESSAGEPARTS(nC) ' get the full message
            Else ' enough parts
                If strPARTSOFMESSAGEPART(0) = "" Then ' if colour not included
                    lCOLOUR = vbBlack ' default: black
                Else ' colour is included
                    lCOLOUR = CLng(strPARTSOFMESSAGEPART(0)) ' get the colour
                End If
                bBOLD = CBool(strPARTSOFMESSAGEPART(1)) ' get the boldness
                strONEMESSAGEPART = strPARTSOFMESSAGEPART(2) ' get the message
                strONEMESSAGEPART = Replace(strONEMESSAGEPART, "&tide;", "~") ' unescape ~
                strONEMESSAGEPART = Replace(strONEMESSAGEPART, "&bslash;", "\") ' unescape \
                strONEMESSAGEPART = Replace(strONEMESSAGEPART, "&amp;", "&") ' unescape &
            End If
        Else
            lCOLOUR = vbBlack
            strONEMESSAGEPART = strMESSAGEPARTS(nC)
        End If
        If currentSTATE = "playing" Then ' if playing
            strCHATLOG(UBound(strCHATLOG)) = strCHATLOG(UBound(strCHATLOG)) & strONEMESSAGEPART ' add string without colour info
        Else ' in lobby
            frmLOBBY.rtbCHATLOG.SelStart = Len(frmLOBBY.rtbCHATLOG.Text) ' start selection at end of text
            frmLOBBY.rtbCHATLOG.SelText = strONEMESSAGEPART ' add new message
            frmLOBBY.rtbCHATLOG.SelStart = Len(frmLOBBY.rtbCHATLOG.Text) - Len(strONEMESSAGEPART) ' start selection at start of new message
            frmLOBBY.rtbCHATLOG.SelLength = Len(strONEMESSAGEPART) ' select new message
            frmLOBBY.rtbCHATLOG.SelColor = lCOLOUR ' change colour
            frmLOBBY.rtbCHATLOG.SelBold = bBOLD ' change boldness
        End If
        nC = nC + 1
    Loop
    
    If currentSTATE = "playing" Then ' if playing
        ' cut off end of message if longer then max chars
        If Len(strTEXT) > maxLENGTHOFMSGINGAME Then ' if message is longer then max length for in game
            strCHATLOG(UBound(strCHATLOG)) = Left$(strCHATLOG(UBound(strCHATLOG)), maxLENGTHOFMSGINGAME - 3) & "..." ' cut off message, and add a "..."
        End If
    Else ' in lobby
        frmLOBBY.scrollDOWNCHATLOG ' scroll down chat log to show newest message
    End If
End Sub

Public Sub handleREQUEST(lARRAYID As Long, strCOMMAND As String, strDESCRIPTION As String)
    Dim nC As Integer ' counter to use in select
    nC = 0
    
    Select Case strCOMMAND
        Case "DISCONNECT" ' disconnect
            If strDESCRIPTION = "" Then
                MsgBox "Disconnected from host!", vbOKOnly, programNAME ' alert that you were disconnected
            Else
                MsgBox "Disconnected from host: " & strDESCRIPTION, vbOKOnly, programNAME ' alert reason that you were disconnected
            End If
            sckDISCONNECTED 0, False ' disconnect
        Case "VERSION" ' server wants version
            cSERVER(0).connected = True ' set as connected
            cSERVER(0).sendString "VERSION", VERSION ' send version
        Case "login" ' server wants username
            If strDESCRIPTION = "" Then ' if server wants username
                cSERVER(0).sendString "login", strNAME ' send name
            Else ' login success
                frmLOBBY.Show ' show lobby form
                Unload frmNEWGAME ' hide the new game form
                currentSTATE = "lobby" ' currently in lobby
            End If
        Case "playerList" ' player list update
            If strDESCRIPTION <> "" Then ' if received names
                Dim strPLAYERLIST() As String
                strPLAYERLIST = Split(strDESCRIPTION, "~") ' split players
                If UBound(strPLAYERLIST) + 1 <> MAXCLIENTS Then ' bad command
                    Exit Sub ' exit
                End If
                Dim strONEPLAYER() As String ' one player's info, score (long), ready (boolean), "playersnamewithescapedchars: &amp;"
                intPLAYERS = 0 ' reset number of players
                nC = 0
                Do While nC < MAXCLIENTS ' for each player
                    strONEPLAYER = Split(strPLAYERLIST(nC), "\", 3) ' separate score, ready state and name
                    If UBound(strONEPLAYER) = 2 Then ' if player exists
                        intPLAYERS = intPLAYERS + 1 ' one more player
                        ccinfoPLAYERINFO(nC).lLEVELSCORE = CLng(strONEPLAYER(0))
                        ccinfoPLAYERINFO(nC).bREADY = CBool(strONEPLAYER(1))
                        ccinfoPLAYERINFO(nC).strNAME = Replace(strONEPLAYER(2), "&tide;", "~") ' unescape ~
                        ccinfoPLAYERINFO(nC).strNAME = Replace(ccinfoPLAYERINFO(nC).strNAME, "&amp;", "&") ' unescape &
                    Else ' player doesn't exist
                        ccinfoPLAYERINFO(nC).reset
                    End If
                    nC = nC + 1 ' next player
                Loop
                If currentSTATE = "lobby" Or currentSTATE = "lobbyShop" Then ' if in lobby
                    frmLOBBY.updatePLAYERLIST ' update player list in lobby
                End If
            End If
        Case "readyState" ' update on user's ready state
            Dim strSTATEPARTS() As String
            strSTATEPARTS = Split(strDESCRIPTION, "\") ' split description into arrayNumber, readyOrNot
            If UBound(strSTATEPARTS) = 1 Then ' if enough parts
                ccinfoPLAYERINFO(CLng(strSTATEPARTS(0))).bREADY = CBool(strSTATEPARTS(1)) ' copy ready state
            End If
            If currentSTATE = "lobby" Or currentSTATE = "lobbyShop" Then ' if in lobby
                frmLOBBY.updatePLAYERLIST ' update player list in lobby
            End If
        Case "playerScore" ' update on user's level score
            Dim strSCOREPARTS() As String
            strSCOREPARTS = Split(strDESCRIPTION, "\") ' split description into arrayNumber, levelScore
            If UBound(strSCOREPARTS) = 1 Then ' if enough parts
                ccinfoPLAYERINFO(CLng(strSCOREPARTS(0))).lLEVELSCORE = CLng(strSCOREPARTS(1)) ' copy level score
            End If
        Case "disableReadyButton" ' countdown has started, disable ready button
            If currentSTATE = "lobbyShop" Then  ' if in shop
                Unload frmSTORE ' unload shop form
                currentSTATE = "lobby" ' now in lobby
            End If
            If currentSTATE = "lobby" Then  ' if in lobby
                frmLOBBY.cmdREADY.Enabled = False ' disable ready button
                frmLOBBY.cmdTOSTORE.Visible = False ' hide open shop button
            End If
        Case "game" ' game start/stop
            If strDESCRIPTION = "start" Then ' if starting game
                frmATTACK.Show ' show game form
                Unload frmLOBBY ' hide lobby screen
                If currentSTATE = "lobbyShop" Then ' if shop is open
                    Unload frmSTORE ' hide shop form
                End If
                currentSTATE = "playing" ' currently playing
                ' clear game chat
                Do While nC <= UBound(strCHATLOG) ' for each chat log place
                    strCHATLOG(nC) = "" ' clear this chat
                    nC = nC + 1 ' next chat log spot
                Loop
            Else ' if stopping game
                If strDESCRIPTION = "stopLoose" Or strDESCRIPTION = "stopLooseHighscore" Then ' lost game
                    bEXIT = True ' stop playing game
                    If strDESCRIPTION = "stopLoose" Then ' didn't have highest score in round
                        frmLOBBY.Show ' show lobby
                        currentSTATE = "lobby" ' currently in lobby
                        frmLOBBY.cmdTOSTORE.Visible = False ' can't see "To shop" button
                    Else ' user lost, but got highest score for the round
                        frmLOBBY.Show ' show the lobby
                        frmLOBBY.cmdTOSTORE.Visible = True ' show the "To Shop" button in the lobby
                        frmSTORE.Show ' show the shop
                        currentSTATE = "lobbyShop" ' currently in lobby and shop
                    End If
                    frmLOBBY.rtbCHATLOG.Text = "You lost the level!" ' alert user that they lost
                Else ' won game
                    bEXIT = True ' stop playing game
                    If strDESCRIPTION = "stopWin" Then ' didn't have highest score in round
                        frmLOBBY.Show ' show the lobby
                        currentSTATE = "lobby" ' currently in lobby
                    Else
                        frmLOBBY.Show ' show the lobby
                        frmLOBBY.cmdTOSTORE.Visible = True ' show the "To Shop" button in the lobby
                        frmSTORE.Show ' show the shop
                        currentSTATE = "lobbyShop" ' currently in lobby and shop
                    End If
                    frmLOBBY.rtbCHATLOG.Text = "You won the level!" ' alert user that they lost
                End If
                Unload frmATTACK ' hide game form
            End If
        Case "maxHealth" ' max health update
            If strDESCRIPTION <> "" Then ' if not bad command
                lCASTLEMAXHEALTH = CLng(strDESCRIPTION) ' get new max health
            End If
            If currentSTATE = "lobbyShop" Then ' if in the shop
                frmSTORE.updateLABELS ' update labels inside the shop
            End If
        Case "health" ' health update
            If strDESCRIPTION <> "" Then ' if not bad command
                lCASTLECURRENTHEALTH = CLng(strDESCRIPTION) ' get new health
            End If
            If currentSTATE = "lobbyShop" Then ' if in the shop
                frmSTORE.updateLABELS ' update labels inside the shop
            End If
        'Case "moneyLevel" ' level money update
        '    If strDESCRIPTION <> "" Then ' if not bad command
        '        lLEVELMONEY = CLng(strDESCRIPTION) ' update the money for the current level
        '    End If
        Case "moneyTotal" ' money update
            If strDESCRIPTION <> "" Then ' if not bad command
                lMONEY = CLng(strDESCRIPTION) ' update money
            End If
            If currentSTATE = "lobbyShop" Then ' if in the shop
                frmSTORE.updateLABELS ' update labels inside the shop
            End If
        Case "monstersLeft" ' money update
            If strDESCRIPTION <> "" Then ' if not bad command
                lMONSTERSLEFT = CLng(strDESCRIPTION) ' update money
            End If
        Case "flaPower" ' flail power update
            If strDESCRIPTION <> "" Then ' if not bad command
                intFLAILPOWER = CInt(strDESCRIPTION) ' update flail power
            End If
            If currentSTATE = "lobbyShop" Then ' if in the shop
                frmSTORE.updateLABELS ' update labels inside the store
            End If
        Case "flaGoThrough" ' flail go through update
            If strDESCRIPTION <> "" Then ' if not bad command
                intFLAILGOTHROUGH = CInt(strDESCRIPTION) ' update flail go through
            End If
            If currentSTATE = "lobbyShop" Then ' if in the shop
                frmSTORE.updateLABELS ' update labels inside the store
            End If
        Case "flaAmount" ' flail amount update
            If strDESCRIPTION <> "" Then ' if not bad command
                intFLAILAMOUNT = CInt(strDESCRIPTION) ' update flail amount
            End If
            If currentSTATE = "lobbyShop" Then ' if in the shop
                frmSTORE.updateLABELS ' update labels inside the store
            End If
        Case "nextLevel" ' current level update (next level user will go on)
            If strDESCRIPTION <> "" Then ' if not bad command
                lCURRENTLEVEL = CLng(strDESCRIPTION) ' update current level
                If currentSTATE = "lobby" Or currentSTATE = "lobbyShop" Then ' if in lobby
                    frmLOBBY.updateNEXTLEVELLBL ' update next level label to show the correct next level
                End If
            End If
        Case "chat" ' somebody is talking or message from server
            handleCHAT strDESCRIPTION ' handle chat message
        Case "updateMon" ' update monster info (can be new/sync/delete)
            If strDESCRIPTION <> "" Then ' if not bad command
                If currentSTATE = "playing" Then ' if currently playing
                    updateMONSTER strDESCRIPTION ' update the monster
                End If
            End If
        Case "updateFlail" ' update flail info (can be new/sync/delete)
            If strDESCRIPTION <> "" Then ' if not bad command
                If currentSTATE = "playing" Then ' if currently playing
                    updateFLAIL strDESCRIPTION ' update the flail
                End If
            End If
    End Select
End Sub
