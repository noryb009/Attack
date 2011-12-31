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
Dim lCURRENTANIFRAME As Long ' current frame in the animation

Public Property Get currentFRAME() As Long ' get the current frame
    currentFRAME = lCURRENTANIFRAME \ ticksPerFrame ' return the current frame / ticks per frame
End Property

Public Property Let currentFRAME(ByVal newFRAME As Long) ' set the current frame
    If newFRAME > -1 And newFRAME < cmontypeMONSTERINFO(intTYPE).lFRAMES * ticksPerFrame Then ' if in bounds
        lCURRENTANIFRAME = newFRAME * ticksPerFrame ' set current frame to newFrame * ticks per frame to start on (frame / ticksPerFrame)
    Else ' out of bounds
        lCURRENTANIFRAME = 0 ' set to 0
    End If
End Property

Private Sub Class_Initialize() ' on class init
    bACTIVE = False ' monster is not active
End Sub

Public Sub nextFRAME() ' go to the next frame
    lCURRENTANIFRAME = lCURRENTANIFRAME + 1 ' add one frame
    If lCURRENTANIFRAME = cmontypeMONSTERINFO(intTYPE).lFRAMES * ticksPerFrame Then ' if past last frame
        lCURRENTANIFRAME = 0 ' go back to first frame
    End If
End Sub

Public Sub moveMONSTER() ' move monster
    sngX = sngX + sngMOVESPEED * sngMOVINGH ' add distance
    If (sngMOVINGH > 0 And sngX + cmontypeMONSTERINFO(intTYPE).lWIDTH > castleWALLLEFT) Or (sngMOVINGH < 0 And sngX < castleWALLRIGHT) Then ' if attacking castle
        lCASTLECURRENTHEALTH = lCASTLECURRENTHEALTH - cmontypeMONSTERINFO(intTYPE).intATTACKPOWER ' take away health
        intMONSTERSATTACKEDCASTLE = intMONSTERSATTACKEDCASTLE + 1 ' add to the number of monsters which attacked castle
        bACTIVE = False ' disable monster
    End If
End Sub
