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
      Width           =   4575
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

Sub logout()
    cSERVER(0).disconnect
    Unload frmLOBBY
    frmNEWGAME.Show
End Sub

Private Sub cmdLOGOUT_Click()
    logout
End Sub

Sub sendMESSAGE()
    cSERVER(0).sendString "chat", Trim(txtMESSAGE.Text)
    txtMESSAGE.Text = ""
    txtMESSAGE.SetFocus
End Sub

Private Sub cmdREADY_Click()
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

Private Sub txtMESSAGE_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then ' pressed enter
        sendMESSAGE
    End If
End Sub

Private Sub Form_Load()
    currentSTATE = "lobby"
    bREADY = False ' not ready
    cmdREADY.Caption = "Not ready"
    cmdREADY.BackColor = vbRed
End Sub
