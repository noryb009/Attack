VERSION 5.00
Begin VB.Form frmLOBBY 
   Caption         =   "Lobby"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSEND 
      Caption         =   "Send"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txtMESSAGE 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2160
      Width           =   2775
   End
   Begin VB.TextBox txtCHATLOG 
      Height          =   1935
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   4575
   End
   Begin VB.CommandButton cmdLOGOUT 
      Caption         =   "Logout"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   2160
      Width           =   735
   End
End
Attribute VB_Name = "frmLOBBY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
End Sub
