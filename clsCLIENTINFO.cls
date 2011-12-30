VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCLIENTINFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

Public strNAME As String ' player name
Public lLEVELSCORE As Long ' player score for this round
Public bREADY As Boolean ' if ready to start game

Public Sub reset() ' reset (after user disconnects)
    strNAME = "" ' empty name
    lLEVELSCORE = 0 ' no level score
    bREADY = False ' not ready
End Sub

Public Sub afterLevelReset() ' reset (after level won/lost)
    lLEVELSCORE = 0 ' no level score
    bREADY = False ' not ready
End Sub

Private Sub Class_Initialize() ' object init
    reset ' set all vars to default
End Sub
