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
Dim bREADY As Boolean

Sub updatePLAYERLIST()
    On Error GoTo noPlayers ' UBound(strPLAYERLIST) gives error if no names
    lstPLAYERS.Clear ' clear old names
    Dim nC As Integer
    nC = 0
    Do While nC <= UBound(strPLAYERLIST) ' for each name
        lstPLAYERS.AddItem strPLAYERLIST(nC) ' add to the listbox
        nC = nC + 1
    Loop
    Exit Sub
noPlayers:
End Sub

Sub logout()
    cSERVER(0).disconnect
    Unload frmLOBBY
    If currentSTATE = "lobbyShop" Then
        Unload frmSTORE
    End If
End Sub

Private Sub cmdLOGOUT_Click()
    logout
    frmNEWGAME.Show
End Sub

Sub sendMESSAGE()
    cSERVER(0).sendString "chat", Trim(txtMESSAGE.Text)
    txtMESSAGE.Text = ""
    txtMESSAGE.SetFocus
End Sub

Private Sub cmdREADY_Click()
    If lCASTLECURRENTHEALTH = 0 Then
        If lMONEY >= 10 Then
            MsgBox "You don't have any health! You can buy more at the store."
        Else
            MsgBox "You don't have any health! Here's a few gold coins for you to buy some at the store."
            lMONEY = 10
        End If
        Exit Sub
    End If
    
    If bREADY = True Then
        bREADY = False ' not ready
        cmdREADY.Caption = "Not ready"
        cmdREADY.BackColor = vbRed
    Else
        bREADY = True ' ready
        cmdREADY.Caption = "Ready!"
        cmdREADY.BackColor = vbGreen
    End If
    cSERVER(0).sendString "ready", CStr(bREADY) ' send to server that you are ready/not ready
End Sub

Private Sub cmdSEND_Click()
    sendMESSAGE
End Sub

Private Sub cmdTOSTORE_Click()
    frmSTORE.Show
End Sub

Private Sub Form_Terminate()
    logout
End Sub

Private Sub txtMESSAGE_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then ' pressed enter
        sendMESSAGE
    End If
End Sub

Private Sub Form_Load()
    If currentSTATE = "" Then ' if from new game
        currentSTATE = "lobby" ' you are in the lobby
    End If
    bREADY = False ' not ready
    cmdREADY.Caption = "Not ready"
    cmdREADY.BackColor = vbRed
    updatePLAYERLIST ' update player list
End Sub
