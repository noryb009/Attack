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
    If onlineMODE = False Then ' if already messaged
        If bMESSAGE = True Then MsgBox "Disconnected from host!" ' message that you disconnected
        onlineMODE = False ' not online anymore
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
    
    Do While lSPOT > UBound(arrMONSTERS) ' if bigger then current array size
        ReDim Preserve arrMONSTERS(0 To UBound(arrMONSTERS) + 1) ' make current array bigger
        Set arrMONSTERS(UBound(arrMONSTERS)) = New clsMONSTER ' set as new monster
    Loop
    
    ' copy new values
    arrMONSTERS(lSPOT).bACTIVE = CBool(arrstrSTATS(1))
    arrMONSTERS(lSPOT).currentFRAME = 0
    arrMONSTERS(lSPOT).intTYPE = CInt(arrstrSTATS(2))
    arrMONSTERS(lSPOT).sngX = CSng(arrstrSTATS(3))
    arrMONSTERS(lSPOT).sngY = CSng(arrstrSTATS(4))
    arrMONSTERS(lSPOT).sngMOVINGH = CSng(arrstrSTATS(5))
    arrMONSTERS(lSPOT).intHEALTH = CLng(arrstrSTATS(6))
End Sub

Sub syncMONSTERS(strALLMONINFO As String) ' sync all the monsters
    Dim nC As Integer
    nC = 0
    ' deactivate all monsters
    Do While nC <= UBound(arrMONSTERS) ' for each monster
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
    
    Dim lSPOT As Long
    lSPOT = CLng(arrstrSTATS(0)) ' get flail spot in arrFLAILS
    
    Do While lSPOT > UBound(arrFLAILS) ' if bigger then current array size
        ReDim Preserve arrFLAILS(0 To UBound(arrFLAILS) + 1) ' make current array bigger
        Set arrFLAILS(UBound(arrFLAILS)) = New clsFLAIL ' set as a new flail
    Loop
    
    ' copy new values
    arrFLAILS(lSPOT).bACTIVE = CBool(arrstrSTATS(1))
    arrFLAILS(lSPOT).sngX = CSng(arrstrSTATS(2))
    arrFLAILS(lSPOT).sngY = CSng(arrstrSTATS(3))
    arrFLAILS(lSPOT).sngMOVINGV = CSng(arrstrSTATS(4))
    arrFLAILS(lSPOT).sngMOVINGH = CSng(arrstrSTATS(5))
    arrFLAILS(lSPOT).intGOTHROUGH = CInt(arrstrSTATS(6))
    If CBool(arrstrSTATS(7)) = True Then ' if we should clear go through
        arrFLAILS(lSPOT).clearWENTTHROUGH ' clear go through
    End If
End Sub

Sub syncFLAILS(strALLFLAINFO As String) ' sync all the flails from the server
    Dim nC As Integer
    nC = 0
    ' deactivate all flails
    Do While nC <= UBound(arrFLAILS) ' for each flail
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

Public Sub handleREQUEST(lARRAYID As Long, strCOMMAND As String, strDESCRIPTION As String)
    Dim nC As Integer ' counter to use in select
    nC = 0
    
    Select Case strCOMMAND
        Case "DISCONNECT" ' disconnect
            If strDESCRIPTION = "" Then
                MsgBox "Disconnected from host!" ' alert that you were disconnected
            Else
                MsgBox "Disconnected from host: " & strDESCRIPTION ' alert reason that you were disconnected
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
                strPLAYERLIST = Split(strDESCRIPTION, "~") ' split names
                ' intPLAYERS is used as a counter
                intPLAYERS = 0 ' reset number of players
                Do While intPLAYERS <= UBound(strPLAYERLIST) ' for each player
                    ' replace special chars
                    strPLAYERLIST(nC) = Replace(strPLAYERLIST(nC), "&tide;", "~")
                    strPLAYERLIST(nC) = Replace(strPLAYERLIST(nC), "&amp;", "&")
                    intPLAYERS = intPLAYERS + 1 ' next player
                Loop
                If currentSTATE = "lobby" Or currentSTATE = "lobbyShop" Then ' if in lobby
                    frmLOBBY.updatePLAYERLIST ' update player list in lobby
                End If
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
                Unload frmSTORE ' hide store screen
                currentSTATE = "playing" ' currently playing
                ' clear game chat
                Do While nC <= UBound(strCHATLOG) ' for each chat log place
                    strCHATLOG(nC) = "" ' clear this chat
                    nC = nC + 1 ' next chat log spot
                Loop
            Else ' if stopping game
                If strDESCRIPTION = "stopLoose" Or strDESCRIPTION = "stopLooseShop" Then ' lost game
                    bEXIT = True ' stop playing game
                    If strDESCRIPTION = "stopLoose" Then ' didn't have highest score in round
                        MsgBox "You lost!" ' alert user that they lost
                        frmLOBBY.Show ' show lobby
                        currentSTATE = "lobby" ' currently in lobby
                        frmLOBBY.cmdTOSTORE.Visible = False ' can't see "To shop" button
                    Else ' user lost, but got highest score for the round
                        MsgBox "You lost!" & vbCrLf & "You got the round high score, so you get to visit the shop!" ' alert user
                        frmLOBBY.Show ' show the lobby
                        frmLOBBY.cmdTOSTORE.Visible = True ' show the "To Shop" button in the lobby
                        frmSTORE.Show ' show the shop
                        currentSTATE = "lobbyShop" ' currently in lobby and shop
                    End If
                Else ' won game
                    bEXIT = True ' stop playing game
                    If strDESCRIPTION = "stopWin" Then ' didn't have highest score in round
                        MsgBox "You won!" ' alert user that they won
                        frmLOBBY.Show ' show the lobby
                        currentSTATE = "lobby" ' currently in lobby
                    Else
                        MsgBox "You won!" & vbCrLf & "You got the round high score, so you get to visit the shop!" ' alert user
                        frmLOBBY.Show ' show the lobby
                        frmLOBBY.cmdTOSTORE.Visible = True ' show the "To Shop" button in the lobby
                        frmSTORE.Show ' show the shop
                        currentSTATE = "lobbyShop" ' currently in lobby and shop
                    End If
                End If
                Unload frmATTACK ' hide game form
            End If
        Case "maxHealth" ' max health update
            If strDESCRIPTION <> "" Then ' if not bad command
                lCASTLEMAXHEALTH = CLng(strDESCRIPTION) ' get new max health
                If currentSTATE = "lobbyShop" Then ' if in the shop
                    frmSTORE.updateLABELS ' update the shop labels
                End If
            End If
        Case "health" ' health update
            If strDESCRIPTION <> "" Then ' if not bad command
                lCASTLECURRENTHEALTH = CLng(strDESCRIPTION) ' get new health
                If currentSTATE = "lobbyShop" Then ' if in the shop
                    frmSTORE.updateLABELS ' update the shop labels
                End If
            End If
        Case "moneyLevel" ' level money update
            If strDESCRIPTION <> "" Then ' if not bad command
                lLEVELMONEY = CLng(strDESCRIPTION) ' update the money for the current level
            End If
        Case "moneyTotal" ' money update
            If strDESCRIPTION <> "" Then ' if not bad command
                lMONEY = CLng(strDESCRIPTION) ' update money
                If currentSTATE = "lobbyShop" Then ' if in shop
                    frmSTORE.updateLABELS ' update shop labels
                End If
            End If
        Case "flaPower" ' flail power update
            If strDESCRIPTION <> "" Then ' if not bad command
                intFLAILPOWER = CInt(strDESCRIPTION) ' update flail power
                If currentSTATE = "lobbyShop" Then ' if in shop
                    frmSTORE.updateLABELS ' update shop labels
                End If
            End If
        Case "flaGoThrough" ' flail go through update
            If strDESCRIPTION <> "" Then ' if not bad command
                intFLAILGOTHROUGH = CInt(strDESCRIPTION) ' update flail go through
                If currentSTATE = "lobbyShop" Then ' if in shop
                    frmSTORE.updateLABELS ' update shop labels
                End If
            End If
        Case "flaAmount" ' flail amount update
            If strDESCRIPTION <> "" Then ' if not bad command
                intFLAILAMOUNT = CInt(strDESCRIPTION) ' update flail amount
                If currentSTATE = "lobbyShop" Then ' if in shop
                    frmSTORE.updateLABELS ' update shop labels
                End If
            End If
        Case "nextLevel" ' current level update (next level user will go on)
            If strDESCRIPTION <> "" Then ' if not bad command
                lCURRENTLEVEL = CLng(strDESCRIPTION) ' update current level
            End If
        Case "chat" ' somebody is talking or message from server
            If currentSTATE = "lobby" Or currentSTATE = "lobbyShop" Then ' in lobby
                If frmLOBBY.txtCHATLOG <> "" Then frmLOBBY.txtCHATLOG = frmLOBBY.txtCHATLOG & vbCrLf ' add newline if not empty
                frmLOBBY.txtCHATLOG = frmLOBBY.txtCHATLOG & strDESCRIPTION ' add to chat log
                frmLOBBY.txtCHATLOG.SelStart = Len(frmLOBBY.txtCHATLOG.Text) ' scroll textbox to show new message
                frmLOBBY.txtCHATLOG.SelLength = 0 ' don't select anything
            ElseIf currentSTATE = "playing" Then ' if playing
                ' bump old messages in chat log
                Do While nC < UBound(strCHATLOG) ' for each (not last) chat log place
                    strCHATLOG(nC) = strCHATLOG(nC + 1) ' move message below up
                    nC = nC + 1 ' next chat log spot
                Loop
                ' add new message
                If Len(strDESCRIPTION) > maxLENGTHOFMSGINGAME Then ' if message is longer then max length for in game
                    strCHATLOG(UBound(strCHATLOG)) = Left$(strDESCRIPTION, maxLENGTHOFMSGINGAME - 3) & "..." ' cut off message, and add a "..."
                Else
                    strCHATLOG(UBound(strCHATLOG)) = strDESCRIPTION ' add full message
                End If
            End If
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
