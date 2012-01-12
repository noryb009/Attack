VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLOBBY 
   Caption         =   "Lobby"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   216
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   409
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtbPLAYERS 
      Height          =   1935
      Left            =   4320
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   3413
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   1
      TextRTF         =   $"frmLOBBY.frx":0000
   End
   Begin RichTextLib.RichTextBox rtbCHATLOG 
      Height          =   1935
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   240
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3413
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmLOBBY.frx":0082
   End
   Begin VB.CommandButton cmdREADY 
      Caption         =   "cmdREADY"
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdLOGOUT 
      Caption         =   "Logout"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdTOSTORE 
      Caption         =   "Open store"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdSEND 
      Caption         =   "Send"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtMESSAGE 
      Height          =   375
      Left            =   0
      MaxLength       =   500
      TabIndex        =   0
      Top             =   2280
      Width           =   5175
   End
   Begin VB.Label lblNEXTLEVEL 
      Caption         =   "lblNEXTLEVEL"
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lblWELCOME 
      Caption         =   "lblWELCOME"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "frmLOBBY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

Dim bREADY As Boolean

Sub updatePLAYERLIST()
    Dim nC As Long
    
    ' sort names alphabeticly
    Dim lORDEROFNAMES(0 To MAXCLIENTS - 1) As Long
    nC = 0
    Do While nC < MAXCLIENTS ' for each name spot
        lORDEROFNAMES(nC) = nC ' count up
        nC = nC + 1
    Loop
    
    Dim lTEMP As Long
    nC = 0
    Do While nC < MAXCLIENTS - 1 ' for each spot (not last spot)
        If (ccinfoPLAYERINFO(lORDEROFNAMES(nC)).strNAME = "" And ccinfoPLAYERINFO(lORDEROFNAMES(nC + 1)).strNAME <> "") Or _
        (ccinfoPLAYERINFO(lORDEROFNAMES(nC + 1)).strNAME <> "" And _
        ccinfoPLAYERINFO(lORDEROFNAMES(nC)).strNAME > ccinfoPLAYERINFO(lORDEROFNAMES(nC + 1)).strNAME) Then  ' if spots should be switched (current is nothing but next has name, or (next is not nothing and next is before current alphabeticly))
            ' switch spots
            lTEMP = lORDEROFNAMES(nC) ' save spot
            lORDEROFNAMES(nC) = lORDEROFNAMES(nC + 1) ' move spot
            lORDEROFNAMES(nC + 1) = lTEMP ' restore spot
            If nC > 0 Then ' if not at first spot
                nC = nC - 1 ' go to previous spot
            End If
        Else
            nC = nC + 1 ' go to next spot
        End If
    Loop
    
    rtbPLAYERS.Text = "" ' clear old names
    nC = 0
    Do While nC < MAXCLIENTS ' for each player
        If ccinfoPLAYERINFO(lORDEROFNAMES(nC)).strNAME <> "" Then ' if player spot is used
            rtbPLAYERS.SelStart = Len(rtbPLAYERS.Text) ' go to end of textbox
            If rtbPLAYERS.Text = "" Then ' if first name
                rtbPLAYERS.SelText = ccinfoPLAYERINFO(lORDEROFNAMES(nC)).strNAME ' add the name
            Else ' not first name
                rtbPLAYERS.SelText = vbCrLf & ccinfoPLAYERINFO(lORDEROFNAMES(nC)).strNAME ' add newline and the name
            End If
            rtbPLAYERS.SelStart = Len(rtbPLAYERS.Text) - Len(ccinfoPLAYERINFO(lORDEROFNAMES(nC)).strNAME) ' start before name that was just added
            rtbPLAYERS.SelLength = Len(ccinfoPLAYERINFO(lORDEROFNAMES(nC)).strNAME) ' select name that was just added
            rtbPLAYERS.SelColor = lPLAYERCOLOURS(lORDEROFNAMES(nC)) ' set colour to player colour
            rtbPLAYERS.SelBold = ccinfoPLAYERINFO(lORDEROFNAMES(nC)).bREADY ' set name to bold if player is ready
        End If
        nC = nC + 1 ' next player
    Loop
End Sub

Sub updateNEXTLEVELLBL() ' update the next level label to show the correct next level
    lblNEXTLEVEL = "Next level: Level " & (lCURRENTLEVEL + 1)
End Sub

Sub logout()
    cSERVER(0).disconnect ' disconnect from server
End Sub

Private Sub cmdLOGOUT_Click()
    onlineMODE = False ' don't alert that you were disconnected from the host
    logout ' logout
    frmNEWGAME.Show ' show new game form
End Sub

Sub sendMESSAGE()
    If txtMESSAGE.Text <> "" Then ' if not an empty message
        cSERVER(0).sendString "chat", Trim(txtMESSAGE.Text) ' send message to server
        txtMESSAGE.Text = "" ' remove message
        txtMESSAGE.SetFocus ' set focus to message box for next message
    End If
End Sub

Public Sub scrollDOWNCHATLOG() ' scroll chat log down to show newest message
    rtbCHATLOG.SelStart = Len(rtbCHATLOG.Text) ' start selection at last char
    rtbCHATLOG.SelLength = 0 ' don't select anything
End Sub

Private Sub cmdREADY_Click()
    If bREADY = True Then ' if was ready, now not
        bREADY = False ' not ready
        cmdREADY.Caption = "Ready" ' change caption
        cmdREADY.BackColor = vbRed ' change back colour of ready button
    Else ' if wasn't ready, now is
        bREADY = True ' ready
        cmdREADY.Caption = "Not ready" ' change caption
        cmdREADY.BackColor = vbGreen ' change back colour of ready button
    End If
    cSERVER(0).sendString "ready", CStr(bREADY) ' send to server that you are ready/not ready
End Sub

Private Sub cmdSEND_Click()
    sendMESSAGE ' send the message to the server
End Sub

Private Sub cmdTOSTORE_Click()
    frmSTORE.Show ' show the store form
End Sub

Private Sub Form_Resize()
    If frmLOBBY.ScaleWidth > 200 And frmLOBBY.ScaleHeight > 120 Then ' if form can fit all the controls
        ' chat log textbox
        rtbCHATLOG.width = frmLOBBY.ScaleWidth - rtbPLAYERS.width
        rtbCHATLOG.height = frmLOBBY.ScaleHeight - rtbCHATLOG.Top - 71 ' alloww room for bottom controls
        ' player list box
        rtbPLAYERS.Left = frmLOBBY.ScaleWidth - rtbPLAYERS.width
        rtbPLAYERS.height = rtbCHATLOG.Top + rtbCHATLOG.height - rtbCHATLOG.Top
        ' new message textbox
        txtMESSAGE.Top = rtbCHATLOG.Top + rtbCHATLOG.height + 7
        txtMESSAGE.width = frmLOBBY.ScaleWidth - cmdSEND.width - 10
        ' send button
        cmdSEND.Top = txtMESSAGE.Top
        cmdSEND.Left = frmLOBBY.ScaleWidth - cmdSEND.width - 3
        ' ready button
        cmdREADY.Top = cmdSEND.Top + cmdSEND.height + 7
        cmdREADY.Left = frmLOBBY.ScaleWidth - cmdREADY.width - 3
        ' logout button
        cmdLOGOUT.Top = cmdREADY.Top
        cmdLOGOUT.Left = cmdREADY.Left - cmdLOGOUT.width - 3
        ' open store button
        cmdTOSTORE.Top = cmdREADY.Top
        cmdTOSTORE.Left = cmdLOGOUT.Left - cmdTOSTORE.width - 3
    End If
End Sub

Private Sub Form_Terminate()
    cSERVER(0).connected = False ' don't alert that you have been disconnected
    logout ' logout from the server
End Sub

Private Sub txtMESSAGE_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then ' pressed enter
        sendMESSAGE ' send the message to the server
    End If
End Sub

Private Sub Form_Load()
    If lCASTLECURRENTHEALTH < 0 Then ' if you are out of health
        lCASTLECURRENTHEALTH = 0 ' reset health
    End If
    If onlineMODE = True And currentSTATE = "" Then ' if from new game
        currentSTATE = "lobby" ' you are in the lobby
    End If
    bREADY = False ' not ready
    cmdREADY.Caption = "Ready" ' change caption
    cmdREADY.BackColor = vbRed ' change back colour of ready button
    updatePLAYERLIST ' update player list
    cmdREADY.Enabled = True ' you can change your ready state
    lblWELCOME.Caption = "Welcome, " & strNAME & "!" ' show your name
    updateNEXTLEVELLBL ' update the next level label to show the next level
    
    Dim nC As Integer
    nC = 0
    ' reset player info
    Do While nC < MAXCLIENTS ' for each player
        ccinfoPLAYERINFO(nC).afterLevelReset ' reset score and ready status
        nC = nC + 1 ' next player
    Loop
    updatePLAYERLIST
End Sub
