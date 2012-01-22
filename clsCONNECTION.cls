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
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

Dim lARRAYID As Long ' this object's number in array of clsCONNECTIONs
Dim bCONNECTED As Boolean ' connected to server
Dim WithEvents sockCONNECTION As Winsock ' connection winsock
Attribute sockCONNECTION.VB_VarHelpID = -1
Dim strBUFFER As String ' stores incomplete received string parts

Public Property Get connected() As Boolean ' get if sockCONNECTION is connected
    connected = bCONNECTED ' return if connected
End Property

Public Property Let connected(bNEWCONNECTED As Boolean) ' set connected to true/false
    bCONNECTED = bNEWCONNECTED ' set bCONNECTED to new value
End Property

Public Property Get ip() As String ' get the IP sockCONNECTION is connected to
    ip = sockCONNECTION.RemoteHostIP ' return remote host ip
End Property

Public Property Let arrayID(newID As Long) ' set the array id
    lARRAYID = newID ' set lARRAYID to the new value
End Property

Sub sendString(strCOMMAND As String, Optional ByVal strTOSEND As String = "") ' send string to sockCONNECTION's connection
    ' replace special chars in command
    strCOMMAND = Replace(strCOMMAND, "&", "&amp;")
    strCOMMAND = Replace(strCOMMAND, "|", "&pipe;")
    strCOMMAND = Replace(strCOMMAND, " ", "&nbsp;")
    
    ' replace special chars in description
    strTOSEND = Replace(strTOSEND, "&", "&amp;")
    strTOSEND = Replace(strTOSEND, "|", "&pipe;")
    strTOSEND = Replace(strTOSEND, " ", "&nbsp;")
    If sockCONNECTION.State = sckConnected Then ' final check if still connected
        sockCONNECTION.SendData strCOMMAND & " " & strTOSEND & "|" ' send the command, description and a pipe (to show where a command starts/ends)
    End If
End Sub

Function connectTOSERVER(strCONNECTTOIP As String) As Boolean
    connectTOSERVER = False ' hasn't succeeded yet
    
    strBUFFER = "" ' clear buffer
    
    If strCONNECTTOIP = "" Then ' if nothing entered
        MsgBox "Please input a valid IP address.", vbOKOnly, programNAME ' alert user
        Exit Function
    End If
    
    Dim longPORT As String ' hold port
    
    Dim strPARTS() As String ' parts of the IP
    ' parse port
    strPARTS = Split(strCONNECTTOIP, ":", 2) ' separate IP and port
    If UBound(strPARTS) < 1 Then ' no port in IP
        longPORT = 23513 ' default port
    Else ' port is included in IP (192.168.0.1:12345)
        longPORT = CLng(strPARTS(1)) ' get the port
        If longPORT < 1024 Or longPORT > 65535 Then ' if out of bounds
            MsgBox "Please input a port between 1024 and 65535.", vbOKOnly, programNAME ' alert user
            Exit Function
        End If
    End If
    
    sockCONNECTION.Connect strPARTS(0), longPORT ' connect to client
    
    connectTOSERVER = True ' success
End Function

Sub acceptCONNECTION(requestID As Long)
    sockCONNECTION.Accept requestID ' accept the request
    bCONNECTED = True ' connected
    strBUFFER = "" ' clear buffer
End Sub

Sub disconnect()
    On Error Resume Next ' if there is an error with closing the socket, ignore it
    sockCONNECTION.Close ' close connection
    On Error GoTo 0 ' return to VB6 error handling
    
    If bCONNECTED = True Then ' if connected
        bCONNECTED = False ' not connected anymore
        sckDISCONNECTED lARRAYID ' run event
    End If
End Sub

Private Sub sockCONNECTION_Close()
    bCONNECTED = False ' not connected
    sockCONNECTION.Close ' close connection
    sckDISCONNECTED lARRAYID ' run disconnect events
End Sub

Private Sub sockCONNECTION_DataArrival(ByVal bytesTotal As Long)
    Dim strRECIEVED As String ' recieved data
    sockCONNECTION.GetData strRECIEVED, vbString ' get the string
    strRECIEVED = strBUFFER & strRECIEVED ' add any text left over from the last message
    
    Dim strCOMMANDS() As String
    strCOMMANDS = Split(strRECIEVED, "|") ' winsock combines commands, so separate at |
    
    Dim strCOMMANDPARTS() As String ' parts of the command
    
    Dim nC As Integer
    Dim nC2 As Integer
    Do While nC < UBound(strCOMMANDS) ' skip last one, strRECIEVED has a "|" on the end
        strCOMMANDPARTS = Split(strCOMMANDS(nC), " ", 2) ' split strRECIEVED into command and description
        
        nC2 = 0
        Do While nC2 <= UBound(strCOMMANDPARTS) ' for each part
            ' convert send-safe chars back to real chars
            strCOMMANDPARTS(nC2) = Replace(strCOMMANDPARTS(nC2), "&nbsp;", " ")
            strCOMMANDPARTS(nC2) = Replace(strCOMMANDPARTS(nC2), "&pipe;", "|")
            strCOMMANDPARTS(nC2) = Replace(strCOMMANDPARTS(nC2), "&amp;", "&")
            nC2 = nC2 + 1 ' next part
        Loop
        
        If UBound(strCOMMANDPARTS) > 0 Then ' if description
            handleREQUEST lARRAYID, strCOMMANDPARTS(0), strCOMMANDPARTS(1) ' handle request
        Else
            handleREQUEST lARRAYID, strCOMMANDPARTS(0), "" ' handle request with empty description
        End If
        nC = nC + 1 ' next command
    Loop
    strBUFFER = strCOMMANDS(UBound(strCOMMANDS)) ' store any extra text for the next command
End Sub

Private Sub sockCONNECTION_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    handleError lARRAYID, Description ' handle the error (different for server/client)
End Sub

Private Sub Class_Initialize() ' object init
    bCONNECTED = False ' not connected
    Set sockCONNECTION = New Winsock ' set up winsock
End Sub

Private Sub Class_Terminate() ' object end
    Set sockCONNECTION = Nothing ' delete winsock
End Sub

Private Sub sockCONNECTION_SendComplete() ' when done sending a command
    If bCONNECTED = False Then ' if not connected anymore (disconnected)
        If sockCONNECTION.State = sckConnected Then ' if still connected
            sockCONNECTION.Close ' close connection
            log "Connection closed: " & sockCONNECTION.RemoteHostIP ' log closing
        End If
    End If
End Sub
