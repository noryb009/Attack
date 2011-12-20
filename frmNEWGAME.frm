VERSION 5.00
Begin VB.Form frmNEWGAME 
   Caption         =   "New Game"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   5265
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
      Caption         =   "Start a new multiplayer game"
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
      Left            =   3720
      TabIndex        =   6
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdLOAD 
      Caption         =   "Load a saved game"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtNAME 
      Height          =   375
      Left            =   120
      MaxLength       =   25
      TabIndex        =   2
      Top             =   960
      Width           =   1695
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
      Width           =   2895
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
    lblNAME.Visible = True
    txtNAME.Visible = True
    cmdNEWSPGAME.Visible = True
    lblIP.Visible = False
    txtIP.Visible = False
    cmdNEWMPGAME.Visible = False
    lstSAVES.Visible = True
    cmdLOAD.Visible = True
    cmdDELETE.Visible = True
    
    cmdSINGLE.Visible = False
    cmdMULTI.Visible = True
    
    frmNEWGAME.height = 3450
End Sub

Private Sub cmdSINGLE_Click()
    showSINGLE
End Sub

Private Sub cmdMULTI_Click()
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
    
    frmNEWGAME.height = 2865
End Sub

Private Sub cmdDELETE_Click()
    If lstSAVES.ListIndex = -1 Then
        MsgBox "Please select the save file to delete"
        Exit Sub
    End If
    If MsgBox("Are you sure you want to delete " & escapeQUOTES(lstSAVES.List(lstSAVES.ListIndex)) & "?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    dbSAVEFILES.Execute "DELETE FROM `SaveGames` WHERE `Name`='" & escapeQUOTES(lstSAVES.List(lstSAVES.ListIndex)) & "'"
    
    loadNamesToListbox
End Sub

Private Sub cmdLOAD_Click()
    'If dataSAVEFILES.Recordset.AbsolutePosition = -1 Then
    If lstSAVES.ListIndex = -1 Then
        MsgBox "Please select your name"
        Exit Sub
    End If
    
    Set recsetSAVES = dbSAVEFILES.OpenRecordset("SELECT * FROM `SaveGames` WHERE `Name`='" & escapeQUOTES(lstSAVES.List(lstSAVES.ListIndex)) & "'")
    
    'strNAME = dataSAVEFILES.Recordset.Fields("Name")
    'lLEVEL = dataSAVEFILES.Recordset.Fields("Level")
    'lCASTLECURRENTHEALTH = dataSAVEFILES.Recordset.Fields("CurrentHealth")
    'lCASTLEMAXHEALTH = dataSAVEFILES.Recordset.Fields("MaxHealth")
    'intFLAILPOWER = dataSAVEFILES.Recordset.Fields("FlailPower")
    
    strNAME = recsetSAVES.Fields("Name")
    lMONEY = recsetSAVES.Fields("Money")
    lLEVEL = recsetSAVES.Fields("Level")
    lCASTLEMAXHEALTH = recsetSAVES.Fields("MaxHealth")
    lCASTLECURRENTHEALTH = recsetSAVES.Fields("CurrentHealth")
    intFLAILPOWER = recsetSAVES.Fields("FlailPower")
    intFLAILGOTHROUGH = recsetSAVES.Fields("FlailGoThrough")
    intFLAILAMOUNT = recsetSAVES.Fields("FlailAmount")
    
    Set recsetsavefiles = Nothing
    Set dbSAVEFILES = Nothing
    
    frmLEVELSELECT.Show
    Unload frmNEWGAME
End Sub

Sub newGAME()
    If Trim(txtNAME.Text) = "" Then ' if empty
        MsgBox "Please input a name."
        Exit Sub
    End If
    
    Dim nC As Integer
    
    If lstSAVES.ListCount <> 0 Then ' for each name in list box
        nC = 0
        'Do While dataSAVEFILES.Recordset.EOF <> True
        Do While nC < lstSAVES.ListCount
            'If dataSAVEFILES.Recordset.Fields("Name") = Trim(txtNAME.Text) Then
            If UCase(lstSAVES.List(nC)) = UCase(Trim(txtNAME.Text)) Then
                MsgBox "Name already exists in database!"
                'dataSAVEFILES.Recordset.Move intID
                Exit Sub
            End If
            'dataSAVEFILES.Recordset.MoveNext
            nC = nC + 1
        Loop
        'dataSAVEFILES.Move intID
    End If
    
    strNAME = Trim(txtNAME.Text)
    lMONEY = 0
    lLEVEL = 1
    lCASTLECURRENTHEALTH = 10
    lCASTLEMAXHEALTH = lCASTLECURRENTHEALTH
    intFLAILPOWER = 1
    intFLAILGOTHROUGH = 1
    intFLAILAMOUNT = 1
    
    Set recsetsavefiles = Nothing
    Set dbSAVEFILES = Nothing
    
    frmLEVELSELECT.Show
    frmLEVELSELECT.saveGAME
    Unload frmNEWGAME
End Sub

Private Sub cmdNEWSPGAME_Click()
    onlineMODE = False
    newGAME
End Sub

Private Sub cmdNEWMPGAME_Click()
    If Trim(txtNAME.Text) = "" Then ' if empty
        MsgBox "Please input a name."
        Exit Sub
    End If
    
    onlineMODE = True
    Set cSERVER(0) = New clsCONNECTION
    cSERVER(0).arrayID = 0
    If cSERVER(0).connectTOSERVER(Trim(txtIP.Text)) = False Then ' if parsing error
        Exit Sub
    End If
    
    strNAME = Trim(txtNAME.Text)
    lMONEY = 0
    lLEVEL = 1
    lCASTLECURRENTHEALTH = 10
    lCASTLEMAXHEALTH = lCASTLECURRENTHEALTH
    intFLAILPOWER = 1
    intFLAILGOTHROUGH = 1
    intFLAILAMOUNT = 1
    
    'cmdMULTI.Enabled = False
    'cmdNEWMPGAME.Enabled = False
End Sub

Sub loadNamesToListbox()
    lstSAVES.Clear
    
    Set recsetSAVES = dbSAVEFILES.OpenRecordset("SELECT `Name` FROM `SaveGames` ORDER BY `Name`")

    If recsetSAVES.RecordCount = 0 Then
        lstSAVES.Visible = False
        cmdLOAD.Visible = False
        frmNEWGAME.height = 1725
    Else
        recsetSAVES.MoveFirst
        Do While recsetSAVES.EOF = False
            lstSAVES.AddItem recsetSAVES.Fields("Name")
            recsetSAVES.MoveNext
        Loop
    End If
End Sub

Private Sub Form_Activate()
    showSINGLE
End Sub

Private Sub Form_Load()
    Set dbSAVEFILES = OpenDatabase(App.Path & "\saveFiles.mdb")
    
    loadNamesToListbox
End Sub

Private Sub txtNAME_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13 ' enter
        newGAME
        KeyAscii = 0
End Select
End Sub
