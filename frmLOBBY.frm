VERSION 5.00
Begin VB.Form frmLOBBY 
   Caption         =   "Lobby"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstPLAYERS 
      Height          =   2010
      Left            =   3720
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdTOSTORE 
      Caption         =   "Open store"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdREADY 
      Caption         =   "cmdREADY"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdSEND 
      Caption         =   "Send"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtMESSAGE 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox txtCHATLOG 
      Height          =   1935
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton cmdLOGOUT 
      Caption         =   "Logout"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   2400
      Width           =   735
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
        cmdREADY.Caption = "Not ready" ' change caption
        cmdREADY.BackColor = vbRed ' change back colour of ready button
    Else ' if wasn't ready, now is
        bREADY = True ' ready
        cmdREADY.Caption = "Ready!" ' change caption
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

Private Sub Form_Terminate()
    logout ' logout from the server
End Sub

Private Sub txtMESSAGE_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then ' pressed enter
        sendMESSAGE ' send the message to the server
    End If
End Sub

Private Sub Form_Load()
    If currentSTATE = "" Then ' if from new game
        currentSTATE = "lobby" ' you are in the lobby
    End If
    bREADY = False ' not ready
    cmdREADY.Caption = "Not ready" ' change caption
    cmdREADY.BackColor = vbRed ' change back colour of ready button
    updatePLAYERLIST ' update player list
End Sub
