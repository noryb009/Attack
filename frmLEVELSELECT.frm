VERSION 5.00
Begin VB.Form frmLEVELSELECT 
   Caption         =   "Form1"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   171
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   303
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "10"
      Height          =   375
      Index           =   9
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "9"
      Height          =   375
      Index           =   8
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "8"
      Height          =   375
      Index           =   7
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "7"
      Height          =   375
      Index           =   6
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "6"
      Height          =   375
      Index           =   5
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "5"
      Height          =   375
      Index           =   4
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "4"
      Height          =   375
      Index           =   3
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdLOGOUT 
      Caption         =   "Logout"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "cmdSAVE"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "3"
      Height          =   375
      Index           =   2
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdSHOP 
      Caption         =   "Visit the shop"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "2"
      Height          =   375
      Index           =   1
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "1"
      Height          =   375
      Index           =   0
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblYOUARE 
      Caption         =   "lblYOUARE"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Choose a level:"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "frmLEVELSELECT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

Private Sub cmdLEVEL_Click(Index As Integer)
    If lCASTLECURRENTHEALTH <= 0 Then
        If lMONEY >= 10 Then
            MsgBox "You don't have any health! You can buy more at the store."
        Else
            MsgBox "You don't have any health! Here's a few gold coins for you to buy some at the store."
            lMONEY = 10
        End If
        Exit Sub
    End If
    
    lCURRENTLEVEL = Index + 1
    
    'default of 0
    Dim nC As Integer
    nC = 0
    Do While nC <= UBound(intMONSTERSONLEVEL)
        intMONSTERSONLEVEL(nC) = 0
        nC = nC + 1
    Loop
    
    Select Case Index + 1
        Case 1
            intMONSTERSONLEVEL(greenMonster) = 10
        Case 2
            intMONSTERSONLEVEL(greenMonster) = 20
            intMONSTERSONLEVEL(blackMonster) = 5
        Case 3
            intMONSTERSONLEVEL(greenMonster) = 10
            intMONSTERSONLEVEL(blackMonster) = 15
            intMONSTERSONLEVEL(bat) = 5
        Case 4
            intMONSTERSONLEVEL(bat) = 25
            intMONSTERSONLEVEL(cloud) = 1
        Case 5
            intMONSTERSONLEVEL(tree) = 25
        Case 6
            intMONSTERSONLEVEL(knightSword) = 15
            intMONSTERSONLEVEL(knightFlail) = 5
        Case 7
            intMONSTERSONLEVEL(cloud) = 5
            intMONSTERSONLEVEL(knightSword) = 25
            intMONSTERSONLEVEL(knightFlail) = 20
            intMONSTERSONLEVEL(knightHorse) = 10
        Case 8
            intMONSTERSONLEVEL(greenMonster) = 20
            intMONSTERSONLEVEL(blackMonster) = 20
            intMONSTERSONLEVEL(bat) = 20
            intMONSTERSONLEVEL(cloud) = 20
            intMONSTERSONLEVEL(knightSword) = 20
            intMONSTERSONLEVEL(knightFlail) = 20
            intMONSTERSONLEVEL(knightHorse) = 15
        Case 9
            intMONSTERSONLEVEL(knightSword) = 20
            intMONSTERSONLEVEL(knightFlail) = 20
            intMONSTERSONLEVEL(knightHorse) = 15
            intMONSTERSONLEVEL(dragon) = 1
        Case 10
            intMONSTERSONLEVEL(dragon) = 15
    End Select
    frmATTACK.Show
    Unload frmLEVELSELECT
End Sub

Private Sub cmdLOGOUT_Click()
    If cmdSAVE.Enabled = True Then
        If MsgBox("Do you want to save?", vbYesNo) = vbYes Then
            saveGAME
        End If
    End If
    frmNEWGAME.Show
    Unload frmLEVELSELECT
End Sub

Sub saveGAME()
    Dim dbSAVEFILES As Database
    Dim recsetSAVES As Recordset
    
    ' open database
    Set dbSAVEFILES = OpenDatabase(App.Path & "\saveFiles.mdb")
    
    Set recsetSAVES = dbSAVEFILES.OpenRecordset("SELECT * FROM `SaveGames` WHERE `Name`='" & escapeQUOTES(strNAME) & "'")
    
    If recsetSAVES.RecordCount = 0 Then
        dbSAVEFILES.Execute "INSERT INTO `SaveGames` (`Name`, `Level`, `MaxHealth`, `CurrentHealth`, `Money`, `FlailGoThrough`, `FlailPower`, `FlailAmount`) VALUES('" & escapeQUOTES(strNAME) & "', '" & lLEVEL & "', '" & lCASTLEMAXHEALTH & "', '" & lCASTLECURRENTHEALTH & "', '" & lMONEY & "', '" & intFLAILGOTHROUGH & "', '" & intFLAILPOWER & "', '" & intFLAILAMOUNT & "')"
    Else
        dbSAVEFILES.Execute "UPDATE `SaveGames` SET `Level`=" & lLEVEL & ", `MaxHealth`=" & lCASTLEMAXHEALTH & ", `CurrentHealth`=" & lCASTLECURRENTHEALTH & ", `Money`=" & lMONEY & ", `FlailGoThrough`=" & intFLAILGOTHROUGH & ", `FlailPower`=" & intFLAILPOWER & ", `FlailAmount`=" & intFLAILAMOUNT & " WHERE `Name`='" & escapeQUOTES(strNAME) & "'"
    End If
    
    Set recsetSAVES = Nothing
    Set dbSAVEFILES = Nothing
End Sub

Private Sub cmdSAVE_Click()
    saveGAME
    
    cmdSAVE.Caption = "Game saved!"
    cmdSAVE.Enabled = False
End Sub

Private Sub cmdSHOP_Click()
    frmSTORE.Show
    Unload frmLEVELSELECT
End Sub

Private Sub Form_Load()
    cmdSAVE.Enabled = True
    cmdSAVE.Caption = "Save game"
    
    lblYOUARE.Caption = "Welcome, " & strNAME & "!"
    
    Dim nC As Integer
    nC = 0
    Do While nC < cmdLEVEL.Count
        If nC < lLEVEL Then
            cmdLEVEL(nC).Visible = True
            cmdLEVEL(nC).BackColor = vbGreen
            If nC = lLEVEL - 1 Then cmdLEVEL(nC).BackColor = vbRed
        Else
            cmdLEVEL(nC).Visible = False
        End If
        nC = nC + 1
    Loop
End Sub
