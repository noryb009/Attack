VERSION 5.00
Begin VB.Form frmSERVER 
   Caption         =   "Attack Server"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6270
   Icon            =   "frmSERVER.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdKICK 
      Caption         =   "Kick"
      Height          =   255
      Left            =   4920
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtKICK 
      Height          =   285
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer timerSYNC 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5280
      Top             =   480
   End
   Begin VB.TextBox txtLOG 
      Height          =   3135
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   480
      Width           =   6255
   End
   Begin VB.CommandButton cmdSTOP 
      Caption         =   "Stop"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdSTART 
      Caption         =   "Start"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtPORT 
      Height          =   285
      Left            =   720
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "23513"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblPLAYERNAME 
      Caption         =   "Name:"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblPORT 
      Caption         =   "Port:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmSERVER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

Dim WithEvents sockLISTEN As Winsock ' winsock to listen for new connections
Attribute sockLISTEN.VB_VarHelpID = -1

Sub showSTART()
    ' hide stop button
    cmdSTOP.Visible = False
    lblPLAYERNAME.Visible = False
    txtKICK.Visible = False
    cmdKICK.Visible = False
    lblPORT.Caption = "Port:"
    ' show start button and port select
    txtPORT.Visible = True
    cmdSTART.Visible = True
End Sub

Sub showSTOP()
    ' show stop button
    cmdSTOP.Visible = True
    lblPLAYERNAME.Visible = True
    txtKICK.Visible = True
    cmdKICK.Visible = True
    ' hide start button and port select
    txtPORT.Visible = False
    cmdSTART.Visible = False
End Sub

Private Sub cmdKICK_Click()
    Dim nC As Long
    nC = 0
    Do While nC < MAXCLIENTS
        If cCLIENTS(nC).connected = True Then
            If UCase(cCLIENTINFO(nC).strNAME) = UCase(txtKICK.Text) Then
                'cCLIENTS(nC).disconnect
                log "You kicked " & cCLIENTINFO(nC).strNAME & "."
                cCLIENTS(nC).connected = False ' not connected any more
                cCLIENTS(nC).sendString "DISCONNECT", "You were kicked from the server."
                broadcast "chat", formatCHATMSG("[") & formatCHATMSG(cCLIENTINFO(nC).strNAME, lPLAYERCOLOURS(nC)) & formatCHATMSG(" was kicked]") ' alert users about kick
                'sckDISCONNECTED nC, True ' somebody disconnected event
            End If
        End If
        nC = nC + 1
    Loop
    broadcastPLAYERLIST ' update user's list of players
    checkIFEVERYONEREADY ' check if everyone is ready to start game
End Sub

Private Sub cmdSTART_Click() ' start the server
    Dim lPORT As Long
    lPORT = Val(txtPORT.Text) ' get the port
    If lPORT < 1024 Or lPORT > 65535 Then ' if out of bounds
        MsgBox "Please input a port between 1024 and 65535", vbOKOnly, programNAME ' alert user
        Exit Sub
    End If
    
    sockLISTEN.LocalPort = lPORT ' listen on port
    
    On Error GoTo couldNotListen ' if port used, send error
    sockLISTEN.Listen ' listen
    
    lblPORT.Caption = "IP: " & sockLISTEN.LocalIP & ":" & sockLISTEN.LocalPort
    
    intPLAYERS = 0 ' 0 players
    lCURRENTLEVEL = 0 ' starting level
    lCASTLECURRENTHEALTH = 10 ' starting health
    lCASTLEMAXHEALTH = lCASTLECURRENTHEALTH ' starting max health
    intFLAILPOWER = 1 ' starting flail power
    intFLAILGOTHROUGH = 1 ' starting amount go through
    intFLAILAMOUNT = 1 ' starting flail amount
    showSTOP ' show the stop GUI
    
    log "Server started on " & sockLISTEN.LocalIP & " at port " & sockLISTEN.LocalPort & "." ' log the IP and port info
    Exit Sub
couldNotListen:
    log "Could not start server: port busy" ' log that the port is busy
End Sub

Private Sub cmdSTOP_Click()
    showSTART ' show start GUI
    
    bFORCEEXIT = True ' stop game if running
    
    Dim nC As Integer
    nC = 0
    Do While nC < MAXCLIENTS ' for each client
        cCLIENTS(nC).disconnect ' disconnect client
        nC = nC + 1
    Loop
    
    sockLISTEN.Close ' close listening winsock
    
    log "Server stopped." ' log that the server has stopped
End Sub

Private Sub Form_Load()
    showSTART ' show the start GUI
    Set sockLISTEN = New Winsock ' make the winsock listener
End Sub

Private Sub Form_Resize() ' on resize
    If frmSERVER.ScaleWidth > 0 And frmSERVER.ScaleHeight > txtLOG.Top Then ' if form is bigger then log textbox
        txtLOG.Left = 0 ' log textbox on left size
        txtLOG.Width = frmSERVER.ScaleWidth ' log textbox is same size as form
        txtLOG.Height = frmSERVER.ScaleHeight - txtLOG.Top ' set the height to go to the bottom of the form
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bFORCEEXIT = True ' exit game if running
    If sockLISTEN.State = sckListening Then ' if listening
        sockLISTEN.Close ' close the listening winsock
    End If
    Set sockLISTEN = Nothing ' delete the listening winsock
End Sub

Private Sub sockLISTEN_ConnectionRequest(ByVal requestID As Long)
    Dim nC As Integer
    nC = 0
    
    Dim bACCEPTED As Boolean ' if client has been accepted
    bACCEPTED = False ' not accepted yet
    
    Do While nC < MAXCLIENTS ' for each client spot
        If cCLIENTS(nC).connected = False Then ' if not used
            cCLIENTS(nC).acceptCONNECTION requestID ' accept the new connection
            log "Connection accepted from " & cCLIENTS(nC).ip ' log the accepting
            cCLIENTS(nC).sendString "VERSION" ' ask for version
            bACCEPTED = True ' has accepted the request
            intPLAYERS = intPLAYERS + 1 ' one more player
            Exit Do ' accepted, don't need to keep looking for empty clients
        End If
        nC = nC + 1 ' next client spot
    Loop
    
    If bACCEPTED = False Then ' if not accepted
        log "Connection rejected, clients full." ' log the rejection
    End If
End Sub

Private Sub timerSYNC_Timer()
    If bPLAYING = True Then ' if playing
        'If timerSYNC.Interval Mod 2 = 0 Then ' every other tick
        '    syncMONSTERS ' sync monsters with clients
        '    timerSYNC.Interval = timerSYNC.Interval + 1 ' set interval to odd number
        'Else ' every other tick
        '    syncFLAILS ' sync flails with clients
        '    timerSYNC.Interval = timerSYNC.Interval - 1 ' set interval to even number
        'End If
    End If
End Sub
