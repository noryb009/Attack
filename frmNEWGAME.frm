VERSION 5.00
Begin VB.Form frmNEWGAME 
   Caption         =   "New Game"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMULTI 
      Caption         =   "Multi player"
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdSINGLE 
      Caption         =   "Single player"
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtIP 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Text            =   "127.0.0.1"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdNEWMPGAME 
      Caption         =   "Join a multiplayer game"
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdDELETE 
      Caption         =   "Delete"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdLOAD 
      Caption         =   "Load a saved game"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtNAME 
      Height          =   375
      Left            =   120
      MaxLength       =   15
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdNEWSPGAME 
      Caption         =   "Start a new game"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
   Begin VB.ListBox lstSAVES 
      Height          =   1230
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblIP 
      Caption         =   "IP:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblNAME 
      Caption         =   "Your name:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblATTACK 
      Caption         =   "Attack"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmNEWGAME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!
 
Dim dbSAVEFILES As Database
Dim recsetSAVES As Recordset

Sub showSINGLE()
    ' show singleplayer GUI, hide multiplayer GUI
    lblNAME.Visible = True
    txtNAME.Visible = True
    cmdNEWSPGAME.Visible = True
    lblIP.Visible = False
    txtIP.Visible = False
    cmdNEWMPGAME.Visible = False
    'lstSAVES.Visible = True ' see below
    cmdLOAD.Visible = True
    cmdDELETE.Visible = True
    
    cmdSINGLE.Visible = False
    cmdMULTI.Visible = True
    
    ' show lstSAVES if there are files, otherwise hide
    If lstSAVES.ListCount > 0 Then
        If frmNEWGAME.WindowState = vbNormal Then ' if in a window
            frmNEWGAME.height = 3500 ' resize form height
        End If
        lstSAVES.Visible = True
    Else
        If frmNEWGAME.WindowState = vbNormal Then ' if in a window
            frmNEWGAME.height = 1950 ' resize form height
        End If
        lstSAVES.Visible = False ' hide listbox
    End If
End Sub
Private Sub cmdSINGLE_Click()
    showSINGLE ' show single player GUI
End Sub

Private Sub cmdMULTI_Click()
    ' show multiplayer GUI, hide singleplayer GUI
    lblNAME.Visible = True
    txtNAME.Visible = True
    cmdNEWSPGAME.Visible = False
    lblIP.Visible = True
    txtIP.Visible = True
    cmdNEWMPGAME.Visible = True
    lstSAVES.Visible = False
    cmdLOAD.Visible = False
    cmdDELETE.Visible = False
    
    cmdSINGLE.Visible = True
    cmdMULTI.Visible = False
    
    ' resize form height
    If frmNEWGAME.WindowState = vbNormal Then ' if in a window
        frmNEWGAME.height = 2865 ' resize form
    End If
End Sub

Sub loadNamesToListbox()
    lstSAVES.Clear ' clear save list
    
    Set recsetSAVES = dbSAVEFILES.OpenRecordset("SELECT `Name` FROM `SaveGames` ORDER BY `Name`") ' get savefile list

    If recsetSAVES.RecordCount <> 0 Then ' if records
        recsetSAVES.MoveFirst ' go to the first record
        Do While recsetSAVES.EOF = False ' while not at the last record
            lstSAVES.AddItem recsetSAVES.Fields("Name") ' add the name to the listbox
            recsetSAVES.MoveNext ' next record
        Loop
    End If
    
    showSINGLE ' update GUI, hide listbox if no records
End Sub

Private Sub cmdDELETE_Click()
    If lstSAVES.ListIndex = -1 Then ' if you don't have a save file selected
        MsgBox "Please select the save file to delete" ' alert the user
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to delete " & escapeQUOTES(lstSAVES.List(lstSAVES.ListIndex)) & "?", vbYesNo) = vbYes Then ' double check
        dbSAVEFILES.Execute "DELETE FROM `SaveGames` WHERE `Name`='" & escapeQUOTES(lstSAVES.List(lstSAVES.ListIndex)) & "'" ' delete from database
        loadNamesToListbox ' refresh listbox
    End If
End Sub

Private Sub cmdLOAD_Click()
    onlineMODE = False ' not online
    intPLAYERS = -1 ' not multiplayer
    
    If lstSAVES.ListIndex = -1 Then ' player hasn't selected a name
        MsgBox "Please select your name" ' alert the user
        Exit Sub
    End If
    
    Set recsetSAVES = dbSAVEFILES.OpenRecordset("SELECT * FROM `SaveGames` WHERE `Name`='" & escapeQUOTES(lstSAVES.List(lstSAVES.ListIndex)) & "'") ' load save file
    
    If recsetSAVES.RecordCount = 0 Then ' if couldn't find router
        MsgBox "Error: could not find save file" ' alert the user
        loadNamesToListbox ' refresh listbox
        Exit Sub
    End If
    
    ' load from save file
    strNAME = recsetSAVES.Fields("Name")
    lMONEY = recsetSAVES.Fields("Money")
    lLEVEL = recsetSAVES.Fields("Level")
    lCASTLEMAXHEALTH = recsetSAVES.Fields("MaxHealth")
    lCASTLECURRENTHEALTH = recsetSAVES.Fields("CurrentHealth")
    intFLAILPOWER = recsetSAVES.Fields("FlailPower")
    intFLAILGOTHROUGH = recsetSAVES.Fields("FlailGoThrough")
    intFLAILAMOUNT = recsetSAVES.Fields("FlailAmount")
    
    frmLEVELSELECT.Show ' show level select form
    Unload frmNEWGAME ' hide this form
End Sub

Sub newGAME()
    If Trim(txtNAME.Text) = "" Then ' if name is empty
        MsgBox "Please input a name." ' alert the user
        Exit Sub
    End If
    
    
    If lstSAVES.ListCount <> 0 Then  ' if there are save files
        Dim nC As Integer
        nC = 0
        Do While nC < lstSAVES.ListCount ' for each name in listbox
            If UCase(lstSAVES.List(nC)) = UCase(Trim(txtNAME.Text)) Then ' if name already used
                MsgBox "Name already exists in database!" ' alert user
                Exit Sub
            End If
            nC = nC + 1 ' next name
        Loop
    End If
    
    onlineMODE = False ' not online
    intPLAYERS = -1 ' single player
    strNAME = Trim(txtNAME.Text) ' store name
    lMONEY = 0 ' defaut money
    lLEVEL = 1 ' starting level
    lCASTLECURRENTHEALTH = 10 ' starting health
    lCASTLEMAXHEALTH = lCASTLECURRENTHEALTH ' starting max health
    intFLAILPOWER = 1 ' starting flail power
    intFLAILGOTHROUGH = 1 ' starting flail go through
    intFLAILAMOUNT = 1 ' starting flail amount
    
    frmLEVELSELECT.Show ' show the level select form
    frmLEVELSELECT.saveGAME ' save the game
    Unload frmNEWGAME ' hide this form
End Sub

Private Sub cmdNEWSPGAME_Click()
    newGAME ' start a new game
End Sub

Private Sub cmdNEWMPGAME_Click()
    If Trim(txtNAME.Text) = "" Then ' if name is empty
        MsgBox "Please input a name." ' alert user
        Exit Sub
    End If
    
    onlineMODE = True ' online
    strNAME = Trim(txtNAME.Text) ' save name
    lMONEY = 0 ' default money
    lLEVEL = 1 ' starting level
    lCASTLECURRENTHEALTH = 10 ' starting health
    lCASTLEMAXHEALTH = lCASTLECURRENTHEALTH ' starting max health
    intFLAILPOWER = 1 ' starting flail power
    intFLAILGOTHROUGH = 1 ' starting flail go through
    intFLAILAMOUNT = 1 ' starting flail amount
    
    Set cSERVER(0) = New clsCONNECTION ' make a new clsCONNECTION, will be able to connect to the server
    cSERVER(0).arrayID = 0 ' array spot 0
    If cSERVER(0).connectTOSERVER(Trim(txtIP.Text)) = False Then ' if parsing error
        Exit Sub
    End If
    
    'cmdMULTI.Enabled = False
    'cmdNEWMPGAME.Enabled = False
End Sub

Private Sub txtNAME_KeyPress(KeyAscii As Integer)
    If onlineMODE = False Then ' if not playing online
        Select Case KeyAscii
            Case 13 ' enter
                newGAME ' start a new game
        End Select
    End If
End Sub

Private Sub Form_Activate()
    loadNamesToListbox ' get save files
End Sub

Private Sub Form_Load()
    Set dbSAVEFILES = OpenDatabase(App.Path & "\saveFiles.mdb") ' open database
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set recsetsavefiles = Nothing ' close recordset
    Set dbSAVEFILES = Nothing ' close database
End Sub
