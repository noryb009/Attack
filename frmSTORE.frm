VERSION 5.00
Begin VB.Form frmSTORE 
   Caption         =   "Store"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFLAILAMOUNT 
      Caption         =   "cmdFLAILAMOUNT"
      Height          =   735
      Left            =   3120
      TabIndex        =   14
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdFLAILGOTHROUGH 
      Caption         =   "cmdFLAILGOTHROUGH"
      Height          =   735
      Left            =   3120
      TabIndex        =   13
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdMOREHEALTH 
      Caption         =   "+10000 Max Health - $100000"
      Height          =   495
      Index           =   3
      Left            =   1560
      TabIndex        =   12
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdMOREHEALTH 
      Caption         =   "+1000 Max Health - $10000"
      Height          =   495
      Index           =   2
      Left            =   1560
      TabIndex        =   11
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdMOREHEALTH 
      Caption         =   "+100 Max Health - $1000"
      Height          =   495
      Index           =   1
      Left            =   1560
      TabIndex        =   10
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdMOREHEALTH 
      Caption         =   "+10 Max Health - $100"
      Height          =   495
      Index           =   0
      Left            =   1560
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdFLAILPOWER 
      Caption         =   "cmdFLAILPOWER"
      Height          =   735
      Left            =   3120
      TabIndex        =   8
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdHEALALL 
      Caption         =   "Heal all"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdHEAL 
      Caption         =   "Heal 10000 - $10000"
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdHEAL 
      Caption         =   "Heal 1000 - $1000"
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdHEAL 
      Caption         =   "Heal 100 - $100"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdHEAL 
      Caption         =   "Heal 10 - $10"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdBACK 
      Caption         =   "Back"
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblMONEY 
      Caption         =   "lblMONEY"
      Height          =   255
      Left            =   2880
      TabIndex        =   7
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

Dim lFLAILUPGRADECOSTS(0 To 9) As Long

Sub loadFLAILUPGRADECOSTS()
    lFLAILUPGRADECOSTS(0) = 100 '0    1 -> 2
    lFLAILUPGRADECOSTS(1) = 200 '0    2 -> 3
    lFLAILUPGRADECOSTS(2) = 400 '0    3 -> 4
    lFLAILUPGRADECOSTS(3) = 800 '0    4 -> 5
    lFLAILUPGRADECOSTS(4) = 1000 '0   5 -> 6
    lFLAILUPGRADECOSTS(5) = 1500 '0   6 -> 7
    lFLAILUPGRADECOSTS(6) = 2000 '0   7 -> 8
    lFLAILUPGRADECOSTS(7) = 3000 '0   8 -> 9
End Sub

Sub updateLABELS()
    If lCASTLECURRENTHEALTH <> 0 Then
        lblCURRENTHEALTH.Caption = "Current health: " & lCASTLECURRENTHEALTH & "0/" & lCASTLEMAXHEALTH & "0"
    Else
        lblCURRENTHEALTH.Caption = "Current health: 0/" & lCASTLEMAXHEALTH & "0"
    End If
    If lMONEY <> 0 Then
        lblMONEY = "$" & lMONEY & "0"
    Else
        lblMONEY = "$0"
    End If
    Dim nC As Integer
    nC = 0
    Do While nC < 4
        If (10 ^ (nC + 1)) \ 10 <= lCASTLEMAXHEALTH - lCASTLECURRENTHEALTH And ((10 ^ (nC + 1)) \ 10) <= lMONEY Then
            cmdHEAL(nC).Enabled = True
        Else
            cmdHEAL(nC).Enabled = False
        End If
        cmdHEAL(nC).Caption = "Heal " & (10 ^ (nC + 1)) & " - $" & (10 ^ (nC + 1))
        
        If (10 ^ (nC + 2)) \ 10 <= lMONEY Then
            cmdMOREHEALTH(nC).Enabled = True
        Else
            cmdMOREHEALTH(nC).Enabled = False
        End If
        cmdMOREHEALTH(nC).Caption = "+" & (10 ^ (nC + 1)) & " health - $" & (10 ^ (nC + 2))
        
        nC = nC + 1
    Loop
    
    If lCASTLEMAXHEALTH <> lCASTLECURRENTHEALTH And lMONEY <> 0 Then
        cmdHEALALL.Enabled = True
    Else
        cmdHEALALL.Enabled = False
    End If
    
    If intFLAILPOWER < 10 Then
        cmdFLAILPOWER.Caption = "Increase flail attack power: " & intFLAILPOWER & " => " & intFLAILPOWER + 1 & vbCrLf & "$" & lFLAILUPGRADECOSTS(intFLAILPOWER) & "0"
        If lMONEY >= lFLAILUPGRADECOSTS(intFLAILPOWER) Then
            cmdFLAILPOWER.Enabled = True
        Else
            cmdFLAILPOWER.Enabled = False
        End If
    Else
        cmdFLAILPOWER.Caption = "Flail attack power fully upgraded!"
        cmdFLAILPOWER.Enabled = False
    End If
    
    If intFLAILGOTHROUGH < 10 Then
        cmdFLAILGOTHROUGH.Caption = "Increase flail piercing power: " & intFLAILGOTHROUGH & " => " & intFLAILGOTHROUGH + 1 & vbCrLf & "$" & lFLAILUPGRADECOSTS(intFLAILGOTHROUGH) & "0"
        If lMONEY >= lFLAILUPGRADECOSTS(intFLAILGOTHROUGH) Then
            cmdFLAILGOTHROUGH.Enabled = True
        Else
            cmdFLAILGOTHROUGH.Enabled = False
        End If
    Else
        cmdFLAILGOTHROUGH.Caption = "Flail piercing power fully upgraded!"
        cmdFLAILGOTHROUGH.Enabled = False
    End If
    
    If intFLAILAMOUNT < 10 Then
        cmdFLAILAMOUNT.Caption = "Increase number of flails: " & intFLAILAMOUNT & " => " & intFLAILAMOUNT + 1 & vbCrLf & "$" & lFLAILUPGRADECOSTS(intFLAILAMOUNT) & "0"
        If lMONEY >= lFLAILUPGRADECOSTS(intFLAILAMOUNT) Then
            cmdFLAILAMOUNT.Enabled = True
        Else
            cmdFLAILAMOUNT.Enabled = False
        End If
    Else
        cmdFLAILAMOUNT.Caption = "Number of flails fully upgraded!"
        cmdFLAILAMOUNT.Enabled = False
    End If
End Sub

Private Sub cmdBACK_Click()
    frmLEVELSELECT.Show
    Unload frmSTORE
End Sub

Private Sub cmdFLAILAMOUNT_Click()
    If onlineMODE = True Then
        cSERVER(0).sendString "buy", "amount~" & lFLAILUPGRADECOSTS(intFLAILAMOUNT)
    End If
    lMONEY = lMONEY - lFLAILUPGRADECOSTS(intFLAILAMOUNT)
    intFLAILAMOUNT = intFLAILAMOUNT + 1
    updateLABELS
End Sub

Private Sub cmdFLAILGOTHROUGH_Click()
    If onlineMODE = True Then
        cSERVER(0).sendString "buy", "goThrough~" & lFLAILUPGRADECOSTS(intFLAILGOTHROUGH)
    End If
    lMONEY = lMONEY - lFLAILUPGRADECOSTS(intFLAILGOTHROUGH)
    intFLAILGOTHROUGH = intFLAILGOTHROUGH + 1
    updateLABELS
End Sub

Private Sub cmdFLAILPOWER_Click()
    If onlineMODE = True Then
        cSERVER(0).sendString "buy", "power~" & lFLAILUPGRADECOSTS(intFLAILPOWER)
    End If
    lMONEY = lFLAILUPGRADECOSTS(intFLAILPOWER)
    intFLAILPOWER = intFLAILPOWER + 1
    updateLABELS
End Sub

Private Sub cmdHEAL_Click(Index As Integer)
    Dim lCOST As Long ' because it costs $1(added 0) for 1 health, the cost is the same
    lCOST = (10 ^ (Index + 1)) \ 10
    If onlineMODE = True Then
        cSERVER(0).sendString "heal", CStr(lCOST) & "~" & CStr(lCOST)
    End If
    lMONEY = lMONEY - lCOST
    lCASTLECURRENTHEALTH = lCASTLECURRENTHEALTH + lCOST
    updateLABELS
End Sub

Private Sub cmdHEALALL_Click()
    If lMONEY > (lCASTLEMAXHEALTH - lCASTLECURRENTHEALTH) Then
        If onlineMODE = True Then
            cSERVER(0).sendString "heal", CStr(lCASTLEMAXHEALTH - lCASTLECURRENTHEALTH) & "~" & CStr(lCASTLEMAXHEALTH - lCASTLECURRENTHEALTH)
        End If
        lMONEY = lMONEY - (lCASTLEMAXHEALTH - lCASTLECURRENTHEALTH)
        lCASTLECURRENTHEALTH = lCASTLEMAXHEALTH
    Else
        If onlineMODE = True Then
            cSERVER(0).sendString "heal", CStr(lMONEY) & "~" & CStr(lMONEY)
        End If
        lCASTLECURRENTHEALTH = lCASTLECURRENTHEALTH + lMONEY
        lMONEY = 0
    End If
    updateLABELS
End Sub

Private Sub cmdMOREHEALTH_Click(Index As Integer)
    Dim lCOST As Long
    Dim lHEALTHADDED As Long
    lCOST = ((10 ^ (Index + 2)) \ 10)
    lHEALTHADDED = ((10 ^ (Index + 1)) \ 10)
    
    If onlineMODE = True Then
        cSERVER(0).sendString "addHealth", CStr(lCOST) & "~" & CStr(lHEALTHADDED)
    End If
    
    lMONEY = lMONEY - lCOST
    lCASTLEMAXHEALTH = lCASTLEMAXHEALTH + lHEALTHADDED
    lCASTLECURRENTHEALTH = lCASTLECURRENTHEALTH + lHEALTHADDED
    updateLABELS
End Sub

Private Sub Form_Load()
    loadFLAILUPGRADECOSTS
    If onlineMODE = True Then
        cmdBACK.Visible = False
    End If
    updateLABELS
End Sub

Private Sub Form_Terminate()
    If onlineMODE = True Then
        currentSTATE = "lobby"
        frmLOBBY.Show
    End If
End Sub
