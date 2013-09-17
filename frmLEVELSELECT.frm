VERSION 5.00
Begin VB.Form frmLEVELSELECT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Level Select"
   ClientHeight    =   3000
   ClientLeft      =   -15
   ClientTop       =   285
   ClientWidth     =   4545
   Icon            =   "frmLEVELSELECT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   303
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "Endless"
      Height          =   375
      Index           =   10
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1800
      Width           =   1815
   End
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
      TabStop         =   0   'False
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "cmdSAVE"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2400
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
      Top             =   2400
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
    If lCASTLECURRENTHEALTH <= 0 Then ' if you don't have health
        If lMONEY >= 10 Then ' if you have enough money to buy health
            ' tell user to buy health
            MsgBox "You don't have any health! You can buy more at the store.", vbOKOnly, programNAME
        Else
            ' tell user to buy health, and give some money for them to buy it
            MsgBox "You don't have any health! Here's a few gold coins for you to buy some at the store.", vbOKOnly, programNAME
            lMONEY = 10
        End If
        Exit Sub
    End If
    
    lCURRENTLEVEL = Index + 1 ' set the current level
    
    ' default: 0 monsters of each type
    Dim nC As Integer
    nC = 0
    Do While nC < numberOfMonsters ' for each monster spot
        intMONSTERSONLEVEL(nC) = 0 ' no monsters on this level of this type
        nC = nC + 1
    Loop
    
    Select Case lCURRENTLEVEL ' monsters on current level
        Case 1 ' on level 1
            intMONSTERSONLEVEL(greenMonster) = 10 ' there are 10 green monsters
        Case 2
            intMONSTERSONLEVEL(greenMonster) = 20
            intMONSTERSONLEVEL(blackMonster) = 5
        Case 3
            intMONSTERSONLEVEL(greenMonster) = 10
            intMONSTERSONLEVEL(blackMonster) = 15
            intMONSTERSONLEVEL(bat) = 5
        Case 4
            intMONSTERSONLEVEL(bat) = 25
            intMONSTERSONLEVEL(motherBat) = 1
        Case 5
            intMONSTERSONLEVEL(tree) = 15
            intMONSTERSONLEVEL(goo) = 15
        Case 6
            intMONSTERSONLEVEL(knightSword) = 15
            intMONSTERSONLEVEL(knightFlail) = 5
        Case 7
            intMONSTERSONLEVEL(cloud) = 5
            intMONSTERSONLEVEL(knightSword) = 25
            intMONSTERSONLEVEL(knightFlail) = 20
            intMONSTERSONLEVEL(knightHorse) = 10
        Case 8
            intMONSTERSONLEVEL(greenMonster) = 15
            intMONSTERSONLEVEL(blackMonster) = 15
            intMONSTERSONLEVEL(bat) = 15
            intMONSTERSONLEVEL(cloud) = 15
            intMONSTERSONLEVEL(goo) = 10
            intMONSTERSONLEVEL(knightSword) = 10
            intMONSTERSONLEVEL(knightFlail) = 10
            intMONSTERSONLEVEL(knightHorse) = 5
        Case 9
            intMONSTERSONLEVEL(knightSword) = 20
            intMONSTERSONLEVEL(knightFlail) = 20
            intMONSTERSONLEVEL(knightHorse) = 15
            intMONSTERSONLEVEL(dragon) = 1
        Case 10
            intMONSTERSONLEVEL(motherBat) = 10
            intMONSTERSONLEVEL(dragon) = 15
        Case Else ' not defined above
            generateMONSTERS 10 + (lCURRENTLEVEL * 20) ' generate monsters
    End Select
    frmATTACK.Show ' show the main game form
    Unload frmLEVELSELECT ' hide this form
End Sub

Private Sub cmdLOGOUT_Click()
    If cmdSAVE.Enabled = True Then ' if user hasn't saved yet
        If MsgBox("Do you want to save?", vbYesNo, programNAME) = vbYes Then ' offer to save
            saveGAME ' save game
        End If
    End If
    frmNEWGAME.Show ' show the new game form
    Unload frmLEVELSELECT ' hide this form
End Sub

Sub saveGAME()
    Dim dbSAVEFILES As Database ' database link
    Dim recsetSAVES As Recordset ' record set
    
    Set dbSAVEFILES = OpenDatabase(strDATABASEPATH) ' open database
    
    Set recsetSAVES = dbSAVEFILES.OpenRecordset("SELECT `Name` FROM `SaveGames` WHERE `Name`='" & escapeQUOTES(strNAME) & "'") ' get all rows with current username
    
    If recsetSAVES.RecordCount = 0 Then ' if not inserted yet
        ' insert new save row into the database
        dbSAVEFILES.Execute "INSERT INTO `SaveGames` (`Name`, `Level`, `MaxHealth`, `CurrentHealth`, `Money`, `FlailGoThrough`, `FlailPower`, `FlailAmount`, `Highscore`) VALUES('" & escapeQUOTES(strNAME) & "', '" & lLEVEL & "', '" & lCASTLEMAXHEALTH & "', '" & lCASTLECURRENTHEALTH & "', '" & lMONEY & "', '" & intFLAILGOTHROUGH & "', '" & intFLAILPOWER & "', '" & intFLAILAMOUNT & "', '" & lHIGHSCORE & "')"
    Else
        ' update the save row
        dbSAVEFILES.Execute "UPDATE `SaveGames` SET `Level`=" & lLEVEL & ", `MaxHealth`=" & lCASTLEMAXHEALTH & ", `CurrentHealth`=" & lCASTLECURRENTHEALTH & ", `Money`=" & lMONEY & ", `FlailGoThrough`=" & intFLAILGOTHROUGH & ", `FlailPower`=" & intFLAILPOWER & ", `FlailAmount`=" & intFLAILAMOUNT & ", `Highscore`=" & lHIGHSCORE & " WHERE `Name`='" & escapeQUOTES(strNAME) & "'"
    End If
    
    Set recsetSAVES = Nothing ' close the recordset
    Set dbSAVEFILES = Nothing ' close the database link
End Sub

Private Sub cmdSAVE_Click()
    saveGAME ' save the game
    
    cmdSAVE.Caption = "Game saved!" ' user has saved the game
    cmdSAVE.Enabled = False ' user can't save again
End Sub

Private Sub cmdSHOP_Click()
    frmSTORE.Show ' show the shop
    Unload frmLEVELSELECT ' hide this form
End Sub

Private Sub Form_Load()
    cmdSAVE.Enabled = True ' user can still save
    cmdSAVE.Caption = "Save game" ' user hasn't saved yet
    
    lblYOUARE.Caption = "Welcome, " & strNAME & "!" ' display username
    
    Dim nC As Integer
    nC = 0
    Do While nC < cmdLEVEL.Count ' for each level button
        If nC < lLEVEL Or nC = 10 Then ' if user has beaten level, is the next level, or is endless mode
            cmdLEVEL(nC).Visible = True ' show button
            If nC + 1 < lLEVEL Then ' if not beaten yet
                cmdLEVEL(nC).BackColor = vbGreen ' green background
            Else
                cmdLEVEL(nC).BackColor = vbRed ' red background
            End If
        Else
            cmdLEVEL(nC).Visible = False ' hide level
        End If
        nC = nC + 1 ' next level button
    Loop
End Sub

