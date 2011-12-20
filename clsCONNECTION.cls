VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCONNECTION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim lARRAYID As Long
Dim bCONNECTED As Boolean
Dim WithEvents sockCONNECTION As Winsock
Attribute sockCONNECTION.VB_VarHelpID = -1

Public Property Get connected() As Boolean
    connected = bCONNECTED
End Property

Public Property Let connected(bNEWCONNECTED As Boolean)
    bCONNECTED = bNEWCONNECTED
End Property

Public Property Get ip() As String
    ip = sockCONNECTION.RemoteHostIP
End Property

Public Property Let arrayID(newID As Long)
    lARRAYID = newID
End Property

Sub sendString(strCOMMAND As String, Optional strTOSEND As String = "")
    Replace strCOMMAND, "&", "&amp;"
    Replace strCOMMAND, "|", "&pipe;"
    Replace strCOMMAND, " ", "&nbsp;"
    
    Replace strCOMMAND, "&", "&amp;"
    Replace strCOMMAND, "|", "&pipe;"
    Replace strCOMMAND, " ", "&nbsp;"
    
    sockCONNECTION.SendData strCOMMAND & " " & strTOSEND & "|"
End Sub

Function connectTOSERVER(strCONNECTTOIP As String) As Boolean
    connectTOSERVER = False ' hasn't succeeded yet
    
    If strCONNECTTOIP = "" Then
        MsgBox "Please input a valid IP address."
        Exit Function
    End If
    
    Dim strIP As String ' hold IP
    Dim longPORT As String ' hold port
    
    Dim strPARTS() As String
    ' parse port
    strPARTS = Split(strCONNECTTOIP, ":", 2) ' separate IP and port
    If UBound(strPARTS) < 1 Then ' no port in IP
        longPORT = 23513 ' default port
    Else
        longPORT = CLng(strPARTS(1)) ' get the port
        If longPORT < 1024 Or longPORT > 65535 Then ' if out of bounds
            MsgBox "Please input a port between 1024 and 65535."
            Exit Function
        End If
    End If
    
    strPARTS = Split(strPARTS(0), ".") ' get different IP parts
    If UBound(strPARTS) <> 3 Then ' not enough "."
        MsgBox "Please input a valid IP address."
        Exit Function
    End If
    
    Dim nC As Integer
    nC = 0
    Do While nC < 4 ' for each number between dots
        If CInt(strPARTS(nC)) < 0 Or CInt(strPARTS(nC)) > 255 Then ' if out of bounds
            MsgBox "Please input a valid IP address."
            Exit Function
        End If
        nC = nC + 1
    Loop
    
    strIP = CStr(CLng(strPARTS(0))) & "." & CStr(CLng(strPARTS(1))) & "." & CStr(CLng(strPARTS(2))) & "." & CStr(CLng(strPARTS(3))) ' make new IP
    
    sockCONNECTION.Connect strIP, longPORT
    
    connectTOSERVER = True ' success
End Function

Sub acceptCONNECTION(requestID As Long)
    sockCONNECTION.Accept requestID
    bCONNECTED = True
End Sub

Sub disconnect()
    If bCONNECTED = True Then
        bCONNECTED = False
        sockCONNECTION.Close
        sckDISCONNECTED lARRAYID
    End If
End Sub

Private Sub sockCONNECTION_Close()
    bCONNECTED = False
    sockCONNECTION.Close
    sckDISCONNECTED lARRAYID
End Sub

Private Sub sockCONNECTION_DataArrival(ByVal bytesTotal As Long)
    Dim strRECIEVED As String
    sockCONNECTION.GetData strRECIEVED, vbString
    
    Dim strCOMMANDS() As String
    strCOMMANDS = Split(strRECIEVED, "|")
    
    Dim strCOMMANDPARTS() As String
    
    Dim nC As Integer
    Dim nC2 As Integer
    Do While nC < UBound(strCOMMANDS) ' skip last one, strRECIEVED has a "|" on the end
        strCOMMANDPARTS = Split(strCOMMANDS(nC), " ", 2)
        
        nC2 = 0
        Do While nC2 <= UBound(strCOMMANDPARTS)
            Replace strCOMMANDPARTS(nC2), "&nbsp;", " "
            Replace strCOMMANDPARTS(nC2), "&pipe;", "|"
            Replace strCOMMANDPARTS(nC2), "&amp;", "&"
            nC2 = nC2 + 1
        Loop
        
        If UBound(strCOMMANDPARTS) > 0 Then
            handleREQUEST lARRAYID, strCOMMANDPARTS(0), strCOMMANDPARTS(1)
        Else
            handleREQUEST lARRAYID, strCOMMANDPARTS(0), ""
        End If
        nC = nC + 1
    Loop
End Sub

Private Sub sockCONNECTION_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    handleError lARRAYID, Description
End Sub

Private Sub Class_Initialize()
    bCONNECTED = False
    Set sockCONNECTION = New Winsock
End Sub

Private Sub Class_Terminate()
    Set sockCONNECTION = Nothing
End Sub

Private Sub sockCONNECTION_SendComplete()
    If bCONNECTED = False Then
        If sockCONNECTION.State = sckConnected Then
            sockCONNECTION.Close
            log "Connection closed: " & sockCONNECTION.RemoteHostIP
        End If
    End If
End Sub
