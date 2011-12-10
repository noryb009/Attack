VERSION 5.00
Begin VB.Form frmSTORE 
   Caption         =   "Store"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFLAILPOWER 
      Caption         =   "cmdFLAILPOWER"
      Height          =   735
      Left            =   2760
      TabIndex        =   9
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdHEALALL 
      Caption         =   "Heal all"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdHEAL 
      Caption         =   "Heal 100000 - $100000"
      Height          =   495
      Index           =   4
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdHEAL 
      Caption         =   "Heal 10000 - $100000"
      Height          =   495
      Index           =   3
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdHEAL 
      Caption         =   "Heal 1000 - $10000"
      Height          =   495
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdHEAL 
      Caption         =   "Heal 100 - $1000"
      Height          =   495
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdHEAL 
      Caption         =   "Heal 10 - $100"
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdBACK 
      Caption         =   "Back"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblMONEY 
      Caption         =   "lblMONEY"
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label lblCURRENTHEALTH 
      Caption         =   "lblCURRENTHEALTH "
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "frmSTORE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

Sub updateLABELS()
    lblCURRENTHEALTH.Caption = "Current health: " & lCASTLECURRENTHEALTH & "/" & lCASTLEMAXHEALTH
    lblMONEY = "$" & lMONEY
    Dim nC As Integer
    Do While nC < 5
        If 10 ^ (nC + 1) <= lCASTLEMAXHEALTH - lCASTLECURRENTHEALTH And 10 ^ (nC + 2) <= lMONEY Then
            cmdHEAL(nC).Enabled = True
        Else
            cmdHEAL(nC).Enabled = False
        End If
        nC = nC + 1
    Loop
    
    If lCASTLEMAXHEALTH <> lCASTLECURRENTHEALTH And lMONEY <> 0 Then
        cmdHEALALL.Enabled = True
    Else
        cmdHEALALL.Enabled = False
    End If
    
    If intFLAILPOWER < 10 Then
        cmdFLAILPOWER.Caption = "Increase flail attack power: " & intFLAILPOWER & " => " & intFLAILPOWER + 1 & vbCrLf & "$" & 10 ^ intFLAILPOWER & "0"
        If lMONEY > ((10 ^ intFLAILPOWER) * 10) Then
            cmdFLAILPOWER.Enabled = True
        Else
            cmdFLAILPOWER.Enabled = False
        End If
    Else
        cmdFLAILPOWER.Caption = "Flail attack power fully upgraded!"
        cmdFLAILPOWER.Enabled = False
    End If
End Sub

Private Sub cmdBACK_Click()
    frmLEVELSELECT.Show
    Unload frmSTORE
End Sub

Private Sub cmdFLAILPOWER_Click()
    lMONEY = lMONEY - ((10 ^ intFLAILPOWER) * 10)
    intFLAILPOWER = intFLAILPOWER + 1
    updateLABELS
End Sub

Private Sub cmdHEAL_Click(Index As Integer)
    lMONEY = lMONEY - (10 ^ (Index + 2))
    lCASTLECURRENTHEALTH = lCASTLECURRENTHEALTH + (10 ^ (Index + 1))
    updateLABELS
End Sub

Private Sub cmdHEALALL_Click()
    If lMONEY > (lCASTLEMAXHEALTH - lCASTLECURRENTHEALTH) * 10 Then
        lMONEY = lMONEY - ((lCASTLEMAXHEALTH - lCASTLECURRENTHEALTH) * 10)
        lCASTLECURRENTHEALTH = lCASTLEMAXHEALTH
    Else
        lCASTLECURRENTHEALTH = lCASTLECURRENTHEALTH + lMONEY \ 10
        lMONEY = lMONEY Mod 10
    End If
    updateLABELS
End Sub

Private Sub Form_Load()
    updateLABELS
End Sub
