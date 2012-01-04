VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMONSTERTYPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

Public intPOINTCOST As Integer ' cost in points, used by generateMONSTERS
Public intMAXHEALTH As Integer ' monster health
Public intSTARTINGY As Integer ' starting Y location (-1 is standing on the ground)
Public sngXSPEED As Single ' monster horizontal speed
Public sngYSPEED As Single ' monster vertical speed
Public intATTACKPOWER As Integer ' amount of health to take away when the monster reaches the castle
Public intMONEYADDEDKILL As Integer ' money to give user when you kill the monster
Public intMONEYADDEDHIT As Integer ' money to give user when you hit (but not kill) the monster
Public lWIDTH As Long ' width of the monster's image
Public lHEIGHT As Long ' height of the monster's image
Public lFRAMES As Long ' number of frames
Public strSPECIAL As String ' special action for each monster

Private Sub Class_Initialize() ' object init
    lFRAMES = 1 ' save server from setting, 1 frame
End Sub
