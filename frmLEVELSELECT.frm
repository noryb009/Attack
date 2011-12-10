VERSION 5.00
Begin VB.Form frmLEVELSELECT 
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   224
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "10"
      Height          =   375
      Index           =   9
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "9"
      Height          =   375
      Index           =   8
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "8"
      Height          =   375
      Index           =   7
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "7"
      Height          =   375
      Index           =   6
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "6"
      Height          =   375
      Index           =   5
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "5"
      Height          =   375
      Index           =   4
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "4"
      Height          =   375
      Index           =   3
      Left            =   1440
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
      Left            =   2880
      TabIndex        =   5
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "3"
      Height          =   375
      Index           =   2
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdSHOP 
      Caption         =   "Visit the shop"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "2"
      Height          =   375
      Index           =   1
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "1"
      Height          =   375
      Index           =   0
      Left            =   360
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
    lCURRENTLEVEL = Index + 1
    
    'default of 0
    intMONSTERSONLEVEL(greenMonster) = 0
    intMONSTERSONLEVEL(blackMonster) = 0
    intMONSTERSONLEVEL(bat) = 0
    intMONSTERSONLEVEL(tree) = 0
    intMONSTERSONLEVEL(cloud) = 0
    intMONSTERSONLEVEL(rabbit) = 0
    intMONSTERSONLEVEL(ladybug) = 0
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
            intMONSTERSONLEVEL(rabbit) = 15
            intMONSTERSONLEVEL(ladybug) = 15
            intMONSTERSONLEVEL(tree) = 5
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
        dbSAVEFILES.Execute "INSERT INTO `SaveGames` (`Name`, `Level`, `FlailPower`, `MaxHealth`, `CurrentHealth`, `Money`) VALUES('" & escapeQUOTES(strNAME) & "', '" & lLEVEL & "', '" & intFLAILPOWER & "', '" & lCASTLEMAXHEALTH & "', '" & lCASTLECURRENTHEALTH & "', '" & lMONEY & "')"
    Else
        dbSAVEFILES.Execute "UPDATE `SaveGames` SET `Level`=" & lLEVEL & ", `FlailPower`=" & intFLAILPOWER & ", `MaxHealth`=" & lCASTLEMAXHEALTH & ", `CurrentHealth`=" & lCASTLECURRENTHEALTH & ", `Money`=" & lMONEY & " WHERE `Name`='" & escapeQUOTES(strNAME) & "'"
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
