Attribute VB_Name = "modREQUESTHANDLER"
Sub sckDISCONNECTED(lARRAYID As Long, Optional bMESSAGE As Boolean = True)
    Unload frmLOBBY
    Unload frmATTACK
    Unload frmSTORE
    frmNEWGAME.Show
    If bMESSAGE = True Then MsgBox "Disconnected from host!"
End Sub

Sub handleError(lARRAYID As Long, strDESCRIPTION As String)
    cSERVER(0).disconnect
    'Unload frmLOBBY
    'Unload frmATTACK
    'Unload frmSTORE
    'frmNEWGAME.Show
    'MsgBox "Disconnected from host!"
    
End Sub

'Sub sendMONINFO()
'    Dim strTOSEND As String
'
'    Dim nC As Integer
'    nC = 0
'    Do While nC < numberOfMonsters ' for each monster type
'        If nC > 0 Then strTOSEND = strTOSEND & "\" ' monster separator
'        strTOSEND = strTOSEND & _
'        cmontypeMONSTERINFO(nC).intATTACKPOWER & "~" & _
'        cmontypeMONSTERINFO(nC).intMAXHEALTH & "~" & _
'        cmontypeMONSTERINFO(nC).intMONEYADDEDHIT & "~" & _
'        cmontypeMONSTERINFO(nC).intMONEYADDEDKILL & "~" & _
'        cmontypeMONSTERINFO(nC).intSTARTINGY & "~" & _
'        cmontypeMONSTERINFO(nC).sngSPEED & "~" & _
'        arrcMONSTERPICS(nC).width & "~" & _
'        arrcMONSTERPICS(nC).height
'        nC = nC + 1
'    Loop
'
'    cSERVER(0).sendString "monInfo", strTOSEND ' send collected data
'End Sub

Sub updateMONSTER(strSTATS As String)
    Dim arrstrSTATS() As String
    arrstrSTATS = Split(strSTATS, "~") ' get different data parts
    
    Dim lSPOT As Long
    lSPOT = CLng(arrstrSTATS(0))
    
    Do While lSPOT > UBound(arrMONSTERS) ' if bigger then current array size
        ReDim Preserve arrMONSTERS(0 To UBound(arrMONSTERS) + 1) ' make current array bigger
        Set arrMONSTERS(UBound(arrMONSTERS)) = New clsMONSTER
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

Sub updateFLAIL(strSTATS As String)
    Dim arrstrSTATS() As String
    arrstrSTATS = Split(strSTATS, "~") ' get different data parts
    
    Dim lSPOT As Long
    lSPOT = CLng(arrstrSTATS(0))
    
    Do While lSPOT > UBound(arrFLAILS) ' if bigger then current array size
        ReDim Preserve arrFLAILS(0 To UBound(arrFLAILS) + 1) ' make current array bigger
        Set arrFLAILS(UBound(arrFLAILS)) = New clsFLAIL
    Loop
    
    ' copy new values
    arrFLAILS(lSPOT).bACTIVE = CBool(arrstrSTATS(1))
    arrFLAILS(lSPOT).sngX = CSng(arrstrSTATS(2))
    arrFLAILS(lSPOT).sngY = CSng(arrstrSTATS(3))
    arrFLAILS(lSPOT).sngMOVINGV = CSng(arrstrSTATS(4))
    arrFLAILS(lSPOT).sngMOVINGH = CSng(arrstrSTATS(5))
    arrFLAILS(lSPOT).intGOTHROUGH = CInt(arrstrSTATS(6))
    If CBool(arrstrSTATS(7)) = True Then
        arrFLAILS(lSPOT).clearWENTTHROUGH
    End If
End Sub

Public Sub handleREQUEST(lARRAYID As Long, strCOMMAND As String, strDESCRIPTION As String)
    Select Case strCOMMAND
        Case "DISCONNECT" ' disconnect
            If strDESCRIPTION = "" Then
                MsgBox "Disconnected from host!"
            Else
                MsgBox strDESCRIPTION
            End If
            sckDISCONNECTED 0, False
        Case "VERSION" ' get version
            cSERVER(0).connected = True
            cSERVER(0).sendString "VERSION", VERSION
            ' server accepted request, frmNEWGAME can continue
            frmLOBBY.Show
            Unload frmNEWGAME
'        Case "monInfo" ' server wants monster info
'            sendMONINFO ' send monster info
        Case "login" ' server wants username
            cSERVER(0).sendString "login", strNAME
        Case "game" ' game start/stop
            If strDESCRIPTION = "start" Then
                frmATTACK.Show
                Unload frmLOBBY
            End If
        Case "chat" ' somebody is talking or message from server
            Select Case currentSTATE
                Case "lobby" ' in lobby
                    If frmLOBBY.txtCHATLOG <> "" Then frmLOBBY.txtCHATLOG = frmLOBBY.txtCHATLOG & vbCrLf ' add newline if not empty
                    frmLOBBY.txtCHATLOG = frmLOBBY.txtCHATLOG & strDESCRIPTION ' add to chat log
                    frmLOBBY.txtCHATLOG.SelStart = Len(frmLOBBY.txtCHATLOG.Text) ' scroll textbox to show new message
                    frmLOBBY.txtCHATLOG.SelLength = 0
            End Select
        Case "updateMon" ' update monster info (can be new/sync/delete)
            If strDESCRIPTION = "" Then
                MsgBox "Empty updateMon received from server!" ' oh no!
            Else
                updateMONSTER strDESCRIPTION ' update the monster
            End If
        Case "updateFlail" ' update flail info (can be new/sync/delete)
            If strDESCRIPTION = "" Then
                MsgBox "Empty updateFLAIL received from server!" ' oh no!
            Else
                updateFLAIL strDESCRIPTION ' update the flail
            End If
    End Select
End Sub
