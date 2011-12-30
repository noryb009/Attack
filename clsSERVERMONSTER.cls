VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMONSTER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

Public bACTIVE As Boolean ' if monster is on the screen
Public sngX As Single ' X location
Public sngY As Single ' Y location
Public sngMOVINGH As Single ' moving speed
Public intTYPE As Integer ' monster type
Public intHEALTH As Integer ' monster health

Private Sub Class_Initialize() ' on class init
    bACTIVE = False ' monster is not active
End Sub

Public Sub moveMONSTER() ' move monster
    sngX = sngX + sngMOVESPEED * sngMOVINGH ' add distance
End Sub
