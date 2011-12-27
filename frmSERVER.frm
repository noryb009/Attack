VERSION 5.00
Begin VB.Form frmSERVER 
   Caption         =   "Attack Server"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLOG 
      Height          =   2655
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   480
      Width           =   4695
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
   Begin VB.Label lblPORT 
      Caption         =   "Port:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmSERVER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents sockLISTEN As Winsock
Attribute sockLISTEN.VB_VarHelpID = -1

Sub showSTART()
    cmdSTOP.Visible = False
    lblPORT.Visible = True
    txtPORT.Visible = True
    cmdSTART.Visible = True
End Sub

Sub showSTOP()
    cmdSTOP.Visible = True
    lblPORT.Visible = False
    txtPORT.Visible = False
    cmdSTART.Visible = False
End Sub

Private Sub cmdSTART_Click()
    lPORT = Val(txtPORT.Text)
    If lPORT < 1 Or lPORT > 65535 Then
        MsgBox "Please input a port between 1 and 65535"
        Exit Sub
    End If
    
    sockLISTEN.LocalPort = lPORT
    
    On Error GoTo couldNotListen
    
    sockLISTEN.Listen
    
    intPLAYERS = 0
    lCURRENTLEVEL = 0
    lCASTLECURRENTHEALTH = 10
    lCASTLEMAXHEALTH = lCASTLECURRENTHEALTH
    intFLAILPOWER = 1
    intFLAILGOTHROUGH = 1
    intFLAILAMOUNT = 1
    showSTOP
    
    log "Server started on " & sockLISTEN.LocalIP & " at port " & sockLISTEN.LocalPort & "."
    Exit Sub
couldNotListen:
    log "Could not start server: port busy"
End Sub

Private Sub cmdSTOP_Click()
    showSTART
    
    bFORCEEXIT = True ' stop game if running
    
    Dim nC As Integer
    nC = 0
    Do While nC < MAXCLIENTS
        cCLIENTS(nC).disconnect
        nC = nC + 1
    Loop
    
    sockLISTEN.Close
    
    log "Server stopped."
End Sub

Private Sub Form_Load()
    showSTART
    Set sockLISTEN = New Winsock
End Sub

Private Sub Form_Resize()
    If frmSERVER.ScaleWidth > 0 And frmSERVER.ScaleHeight > txtLOG.Top Then
        txtLOG.Left = 0
        txtLOG.Width = frmSERVER.ScaleWidth
        txtLOG.Height = frmSERVER.ScaleHeight - txtLOG.Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bFORCEEXIT = True ' exit game if running
    If sockLISTEN.State = sckListening Then
        sockLISTEN.Close
    End If
    Set sockLISTEN = Nothing
End Sub

Private Sub sockLISTEN_ConnectionRequest(ByVal requestID As Long)
    Dim nC As Integer
    nC = 0
    Do While nC < MAXCLIENTS
        If cCLIENTS(nC).connected = False Then
            cCLIENTS(nC).acceptCONNECTION requestID
            log "Connection accepted from " & cCLIENTS(nC).ip
            cCLIENTS(nC).sendString "VERSION"
            Exit Do
        End If
        nC = nC + 1
    Loop
End Sub
