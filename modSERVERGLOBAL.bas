Attribute VB_Name = "modGLOBAL"
Global Const VERSION = "0.0.0.1a"
Global Const SERVER = True
Global Const MAXCLIENTS = 4

'Global bHASINFO As Boolean

Global lPORT As Long

Global cCLIENTS(0 To MAXCLIENTS - 1) As New clsCONNECTION
Global cCLIENTINFO(0 To MAXCLIENTS - 1) As New clsCLIENTINFO
Global intCONNECTEDCLIENTS As Integer

Global arrMONSTERS() As clsMONSTER
Global arrFLAILS() As clsFLAIL

Sub log(strNEWLINE As String)
    frmSERVER.txtLOG.Text = strNEWLINE & vbCrLf & frmSERVER.txtLOG.Text
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
    
    frmSERVER.Show
End Sub
