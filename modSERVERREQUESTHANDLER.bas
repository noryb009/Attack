Attribute VB_Name = "modREQUESTHANDLER"
Sub sckDISCONNECTED(lARRAYID As Long)
    log cCLIENTS(lARRAYID).ip & " (" & cCLIENTINFO(lARRAYID).strNAME & ") disconnected."
    cCLIENTINFO(lARRAYID).strNAME = ""
    cCLIENTINFO(lARRAYID).lngSCORE = 0
End Sub

Sub handleError(lARRAYID As Long, strDESCRIPTION As String)
    log "Error from " & lARRAYID & ":" & strDESCRIPTION
End Sub

Sub spawnFLAIL(strSTATS As String)
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
    Loop
    
    If lFLAILSPOT = -1 Then ' no room left in flail array
        ReDim Preserve arrFLAILS(0 To UBound(arrFLAILS) + 1) ' make array 1 bigger
        Set arrFLAILS(UBound(arrFLAILS)) = New clsFLAIL
    End If
    
    arrFLAILS(lFLAILSPOT).bACTIVE = CBool(arrstrSTATS(0))
    arrFLAILS(lFLAILSPOT).sngX = CSng(arrstrSTATS(1))
    arrFLAILS(lFLAILSPOT).sngY = CSng(arrstrSTATS(2))
    arrFLAILS(lFLAILSPOT).sngMOVINGV = CSng(arrstrSTATS(3))
    arrFLAILS(lFLAILSPOT).sngMOVINGH = CSng(arrstrSTATS(4))
    arrFLAILS(lFLAILSPOT).intGOTHROUGH = CInt(arrstrSTATS(5))
    If CBool(arrstrSTATS(6)) = True Then
        arrFLAILS(lFLAILSPOT).clearWENTTHROUGH
    End If
    
    'TODO: finish
    broadcast "newFla", lFLAILSPOT & "~" & strSTATS
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
            
        Case "newFla"
            If strDESCRIPTION = "" Then
                log "Empty newFla received from " & cCLIENTS(lARRAYID).ip
            Else
                log "newFla received from " & cCLIENTS(lARRAYID).ip
                spawnFLAIL strDESCRIPTION
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
        Case Else
            log "Unknown command from " & cCLIENTS(lARRAYID).ip & ": " & strCOMMAND
    End Select
End Sub
