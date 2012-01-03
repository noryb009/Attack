VERSION 5.00
Begin VB.Form frmSTORE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Store"
   ClientHeight    =   3165
   ClientLeft      =   -15
   ClientTop       =   285
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCLOSE 
      Caption         =   "Close"
      Height          =   375
      Left            =   5520
      TabIndex        =   15
      Top             =   2640
      Width           =   1095
   End
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

Dim lFLAILUPGRADECOSTS(0 To 9) As Long ' cost to upgrade flail power, go through, or amount

Sub loadFLAILUPGRADECOSTS() ' sub to load the cost of flail upgrades
    ' all prices are *10, so $100 will display $1000
    lFLAILUPGRADECOSTS(0) = 0 '  0 -> 1 ' already given
    lFLAILUPGRADECOSTS(1) = 100 '  1 -> 2
    lFLAILUPGRADECOSTS(2) = 200 '  2 -> 3
    lFLAILUPGRADECOSTS(3) = 400 '  3 -> 4
    lFLAILUPGRADECOSTS(4) = 1000 '  4 -> 5
    lFLAILUPGRADECOSTS(5) = 2500 ' 5 -> 6
    lFLAILUPGRADECOSTS(6) = 5000 ' 6 -> 7
    lFLAILUPGRADECOSTS(7) = 10000 ' 7 -> 8
    lFLAILUPGRADECOSTS(8) = 50000 ' 8 -> 9
    lFLAILUPGRADECOSTS(9) = 100000 ' 9 -> 10
End Sub

Function healthDISPLAYAMOUNT(intINDEX As Integer) As Long ' get display amount from index
    healthDISPLAYAMOUNT = (10 ^ (intINDEX + 1)) ' get display amount (*10)
End Function

Function healthAMOUNT(intINDEX As Integer) As Long ' get amount from index
    healthAMOUNT = healthDISPLAYAMOUNT(intINDEX) \ 10 ' get amount
End Function

Function healCOST(intINDEX As Integer) As Long ' get cost of healing
    healCOST = healthAMOUNT(intINDEX) ' same price
End Function

Function moreHEALTHCOST(intINDEX As Integer) As Long ' get cost of buying more max health
    moreHEALTHCOST = healthAMOUNT(intINDEX) * 10 ' $10 per amount
End Function

Sub updateLABELS()
    If lCASTLECURRENTHEALTH <> 0 Then ' if you have health
        lblCURRENTHEALTH.Caption = "Current health: " & lCASTLECURRENTHEALTH & "0/" & lCASTLEMAXHEALTH & "0" ' display current health
    Else
        lblCURRENTHEALTH.Caption = "Current health: 0/" & lCASTLEMAXHEALTH & "0" ' display current health as 0, not 00
    End If
    If lMONEY <> 0 Then ' if you have money
        lblMONEY = "$" & lMONEY & "0" ' display money *10
    Else
        lblMONEY = "$0" ' display 0, not 00
    End If
    Dim nC As Integer
    nC = 0
    Do While nC < cmdHEAL.Count ' for each heal and morehealth button
        If healthAMOUNT(nC) <= lCASTLEMAXHEALTH - lCASTLECURRENTHEALTH And healCOST(nC) <= lMONEY Then ' if it won't overfill your health and you have enough money
            cmdHEAL(nC).Enabled = True ' you can buy this much health
        Else
            cmdHEAL(nC).Enabled = False ' you can't buy this much health
        End If
        cmdHEAL(nC).Caption = "Heal " & healthDISPLAYAMOUNT(nC) & " - $" & healCOST(nC) & "0" ' display health amount and cost
        
        If moreHEALTHCOST(nC) <= lMONEY And lCASTLEMAXHEALTH + healthAMOUNT(nC) <= 10000 Then ' if you have enough money to buy morehealth
            cmdMOREHEALTH(nC).Enabled = True ' you can buy this much morehealth
        Else
            cmdMOREHEALTH(nC).Enabled = False ' you can't buy this much morehealth
        End If
        cmdMOREHEALTH(nC).Caption = "+" & healthDISPLAYAMOUNT(nC) & " health - $" & moreHEALTHCOST(nC) & "0" ' display amount and cost
        
        nC = nC + 1 ' next button
    Loop
    
    If lCASTLEMAXHEALTH <> lCASTLECURRENTHEALTH And lMONEY <> 0 Then ' if you have money and you don't have full health
        cmdHEALALL.Enabled = True ' enable heal all button
    Else
        cmdHEALALL.Enabled = False ' disable heal all button, it wouldn't do anything
    End If
    
    If intFLAILPOWER < 10 Then ' if not fully upgraded
        cmdFLAILPOWER.Caption = "Increase flail attack power: " & intFLAILPOWER & " => " & intFLAILPOWER + 1 & vbCrLf & "$" & lFLAILUPGRADECOSTS(intFLAILPOWER) & "0"
        If lMONEY >= lFLAILUPGRADECOSTS(intFLAILPOWER) Then ' if enough money to upgrade
            cmdFLAILPOWER.Enabled = True ' can upgrade
        Else
            cmdFLAILPOWER.Enabled = False ' can't upgrade
        End If
    Else
        cmdFLAILPOWER.Caption = "Flail attack power fully upgraded!" ' fully upgraded
        cmdFLAILPOWER.Enabled = False ' can't upgrade
    End If
    
    If intFLAILGOTHROUGH < 10 Then ' if not fully upgraded
        cmdFLAILGOTHROUGH.Caption = "Increase flail piercing power: " & intFLAILGOTHROUGH & " => " & intFLAILGOTHROUGH + 1 & vbCrLf & "$" & lFLAILUPGRADECOSTS(intFLAILGOTHROUGH) & "0"
        If lMONEY >= lFLAILUPGRADECOSTS(intFLAILGOTHROUGH) Then ' if enough money to upgrade
            cmdFLAILGOTHROUGH.Enabled = True ' can upgrade
        Else
            cmdFLAILGOTHROUGH.Enabled = False ' can't upgrade
        End If
    Else
        cmdFLAILGOTHROUGH.Caption = "Flail piercing power fully upgraded!" ' fully upgraded
        cmdFLAILGOTHROUGH.Enabled = False ' can't upgrade
    End If
    
    If intFLAILAMOUNT < 10 Then ' if not fully upgraded
        cmdFLAILAMOUNT.Caption = "Increase number of flails: " & intFLAILAMOUNT & " => " & intFLAILAMOUNT + 1 & vbCrLf & "$" & lFLAILUPGRADECOSTS(intFLAILAMOUNT) & "0"
        If lMONEY >= lFLAILUPGRADECOSTS(intFLAILAMOUNT) Then ' if enough money to upgrade
            cmdFLAILAMOUNT.Enabled = True ' can upgrade
        Else
            cmdFLAILAMOUNT.Enabled = False ' can't upgrade
        End If
    Else
        cmdFLAILAMOUNT.Caption = "Number of flails fully upgraded!" ' fully upgraded
        cmdFLAILAMOUNT.Enabled = False ' can't upgrade
    End If
End Sub

Private Sub cmdBACK_Click()
    frmLEVELSELECT.Show ' show the level select form
    Unload frmSTORE ' hide this form
End Sub

Private Sub cmdCLOSE_Click()
    If currentSTATE = "lobbyShop" Then ' if currently in lobby and shop
        currentSTATE = "lobby" ' only in lobby
    End If
    Unload frmSTORE ' hide this form
End Sub

Private Sub cmdFLAILPOWER_Click()
    If onlineMODE = True Then ' if playing multiplayer
        cSERVER(0).sendString "buy", "power~" & lFLAILUPGRADECOSTS(intFLAILPOWER) ' update server
    End If
    lMONEY = lMONEY - lFLAILUPGRADECOSTS(intFLAILPOWER) ' spend money
    intFLAILPOWER = intFLAILPOWER + 1 ' upgrade flail power
    updateLABELS
End Sub

Private Sub cmdFLAILGOTHROUGH_Click()
    If onlineMODE = True Then ' if playing multiplayer
        cSERVER(0).sendString "buy", "goThrough~" & lFLAILUPGRADECOSTS(intFLAILGOTHROUGH) ' update server
    End If
    lMONEY = lMONEY - lFLAILUPGRADECOSTS(intFLAILGOTHROUGH) ' spend money
    intFLAILGOTHROUGH = intFLAILGOTHROUGH + 1 ' upgrade flail gothrough
    updateLABELS
End Sub

Private Sub cmdFLAILAMOUNT_Click()
    If onlineMODE = True Then ' if playing multiplayer
        cSERVER(0).sendString "buy", "amount~" & lFLAILUPGRADECOSTS(intFLAILAMOUNT) ' update server
    End If
    lMONEY = lMONEY - lFLAILUPGRADECOSTS(intFLAILAMOUNT) ' spend money
    intFLAILAMOUNT = intFLAILAMOUNT + 1 ' upgrade flail amount
    updateLABELS
End Sub

Private Sub cmdHEALALL_Click()
    If lMONEY > lCASTLEMAXHEALTH - lCASTLECURRENTHEALTH Then ' if more money then missing health
        If onlineMODE = True Then ' if playing multiplayer
            cSERVER(0).sendString "heal", CStr(lCASTLEMAXHEALTH - lCASTLECURRENTHEALTH) & "~" & CStr(lCASTLEMAXHEALTH - lCASTLECURRENTHEALTH) ' update server
        End If
        lMONEY = lMONEY - (lCASTLEMAXHEALTH - lCASTLECURRENTHEALTH) ' take away money
        lCASTLECURRENTHEALTH = lCASTLEMAXHEALTH ' add health
    Else ' more missing health then money
        If onlineMODE = True Then ' if playing multiplayer
            cSERVER(0).sendString "heal", CStr(lMONEY) & "~" & CStr(lMONEY) ' update server
        End If
        lCASTLECURRENTHEALTH = lCASTLECURRENTHEALTH + lMONEY ' add health
        lMONEY = 0 ' 0 money left
    End If
    updateLABELS
End Sub

Private Sub cmdHEAL_Click(Index As Integer)
    If onlineMODE = True Then ' if playing multiplayer
        cSERVER(0).sendString "heal", CStr(healthAMOUNT(Index)) & "~" & CStr(healCOST(Index)) ' update server
    End If
    lMONEY = lMONEY - healCOST(Index) ' take away money
    lCASTLECURRENTHEALTH = lCASTLECURRENTHEALTH + healthAMOUNT(Index) ' add health
    updateLABELS
End Sub

Private Sub cmdMOREHEALTH_Click(Index As Integer)
    If onlineMODE = True Then ' if playing multiplayer
        cSERVER(0).sendString "addHealth", CStr(moreHEALTHCOST(Index)) & "~" & CStr(healthAMOUNT(Index)) ' update server
    End If
    
    lMONEY = lMONEY - moreHEALTHCOST(Index) ' take away money
    lCASTLEMAXHEALTH = lCASTLEMAXHEALTH + healthAMOUNT(Index) ' add max health
    lCASTLECURRENTHEALTH = lCASTLECURRENTHEALTH + healthAMOUNT(Index) ' add health
    updateLABELS
End Sub

Private Sub Form_Load()
    loadFLAILUPGRADECOSTS ' load the cost of flail upgrades into lFLAILUPGRADECOSTS()
    If onlineMODE = True Then ' if multiplayer
        cmdBACK.Visible = False ' can't go to level select form
    Else ' offline
        cmdCLOSE.Visible = False ' don't show close button
    End If
    updateLABELS
End Sub

Private Sub Form_Terminate()
    If onlineMODE = True Then ' if multiplayer
        currentSTATE = "lobby" ' no longer in the shop
        frmLOBBY.Show ' show the lobby
    End If
End Sub
