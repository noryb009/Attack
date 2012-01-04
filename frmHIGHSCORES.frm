VERSION 5.00
Begin VB.Form frmHIGHSCORES 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBACK 
      Caption         =   "Back"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.ListBox lstHIGHSCORES 
      Height          =   2010
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Highscores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmHIGHSCORES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

Public strWHEREISBACK As String

Private Sub cmdBACK_Click()
    If strWHEREISBACK = "levelSelect" Then ' if going to level select
        frmLEVELSELECT.Show ' show the level select form
    Else ' if going to the new game form or not set
        frmNEWGAME.Show ' show the new game form
    End If
    Unload frmHIGHSCORES ' hide this form
End Sub

Private Sub Form_Load()
    lstHIGHSCORES.Clear ' clear text box
    
    Dim dbSAVEFILES As Database ' database link
    Dim recsetHIGHSCORES As Recordset ' record set
    
    Set dbSAVEFILES = OpenDatabase(strDATABASEPATH) ' open database
    
    ' get list of names and scores where the high score is greater then 0 in descending order
    Set recsetHIGHSCORES = dbSAVEFILES.OpenRecordset("SELECT `Name`, `Highscore` FROM `SaveGames` WHERE `Highscore` > 0 ORDER BY `Highscore` DESC")
    
    If recsetHIGHSCORES.RecordCount = 0 Then ' if no savefiles
        lstHIGHSCORES.AddItem "No highscores!" ' alert user that there are no highscores
        Exit Sub ' exit
    End If
    
    Dim lCURRENTPLACE As Long
    lCURRENTPLACE = 1 ' first person is 1st place
    recsetHIGHSCORES.MoveFirst ' go to first highscore save
    
    Do While recsetHIGHSCORES.EOF = False ' while not at end
        lstHIGHSCORES.AddItem lCURRENTPLACE & ". " & recsetHIGHSCORES.Fields("Name") & " - " & recsetHIGHSCORES.Fields("Highscore") & "0" ' add "1. Name - Score"
        lCURRENTPLACE = lCURRENTPLACE + 1 ' next place (was 1st, now 2nd)
        recsetHIGHSCORES.MoveNext ' next high score
    Loop
    
    Set recsetHIGHSCORES = Nothing ' close recordset
    Set dbSAVEFILES = Nothing ' close database
End Sub
