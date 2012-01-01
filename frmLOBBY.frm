VERSION 5.00
Begin VB.Form frmLOBBY 
   Caption         =   "Lobby"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   216
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   364
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdREADY 
      Caption         =   "cmdREADY"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdLOGOUT 
      Caption         =   "Logout"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   2760
      Width           =   735
   End
   Begin VB.ListBox lstPLAYERS 
      Height          =   2205
      Left            =   4320
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdTOSTORE 
      Caption         =   "Open store"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdSEND 
      Caption         =   "Send"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtMESSAGE 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2280
      Width           =   4455
   End
   Begin VB.TextBox txtCHATLOG 
      Height          =   2175
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   4335
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
    If Not Not strPLAYERLIST Then ' if strPLAYERLIST has names
        lstPLAYERS.Clear ' clear old names
        Dim nC As Integer
        nC = 0
        Do While nC <= UBound(strPLAYERLIST) ' for each name
            lstPLAYERS.AddItem strPLAYERLIST(nC) ' add to the listbox
            nC = nC + 1
        Loop
    End If
End Sub

Sub logout()
    cSERVER(0).disconnect ' disconnect from server
    Unload frmLOBBY ' hide this form
    If currentSTATE = "lobbyShop" Then ' if shop is visible
        Unload frmSTORE ' hide shop
    End If
End Sub

Private Sub cmdLOGOUT_Click()
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
    If frmLOBBY.ScaleWidth > 200 And frmLOBBY.ScaleHeight > 80 Then ' if form can fit all the controls
        ' chat log textbox
        txtCHATLOG.width = frmLOBBY.ScaleWidth - lstPLAYERS.width
        txtCHATLOG.height = frmLOBBY.ScaleHeight - 71 ' alloww room for bottom controls
        ' player list box
        lstPLAYERS.Left = frmLOBBY.ScaleWidth - lstPLAYERS.width
        lstPLAYERS.height = txtCHATLOG.height
        ' new message textbox
        txtMESSAGE.Top = txtCHATLOG.height + 7
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
    If currentSTATE = "lobbyShop" Then  ' if in shop
        Unload frmSTORE ' unload shop form
    End If
    logout ' logout from the server
    currentSTATE = "" ' not in anywhere
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
    cmdREADY.Enabled = True
End Sub
