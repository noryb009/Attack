VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFLAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

Public bACTIVE As Boolean ' if flail is on screen
Public lOWNER As Long ' who shot the flail
Public sngX As Single ' X location
Public sngY As Single ' Y location
Public sngMOVINGH As Single ' X moving speed
Public sngMOVINGV As Single ' Y moving speed
Public lCURRENTANIFRAME As Long ' current animation frame
Public intGOTHROUGH As Integer ' number of go through left
Dim longWENTTHROUGH() As Long ' array of monsters flail has gone through
Dim lSIZEOFlongWENTTHROUGH As Long ' size of longWENTTHROUGH to not have to call VB functions

Public Sub clearWENTTHROUGH() ' clear the go through
    lSIZEOFlongWENTTHROUGH = 0 ' reset the size tracker
End Sub

Public Function didGOTHROUGH(intCHECK As Integer) As Boolean ' check if flail went through monster
    Dim nC As Integer
    nC = 1 ' skip first array spot, not used
    didGOTHROUGH = False ' default: didn't go through
    Do While nC <= lSIZEOFlongWENTTHROUGH ' while less then size of longWENTTHROUGH
        If longWENTTHROUGH(nC) = intCHECK Then ' if flail went through this monster
            didGOTHROUGH = True ' did go through monster
            Exit Do ' found monster, don't need to keep looking
        End If
        nC = nC + 1 ' next go through
    Loop
End Function

Public Sub addGOTHROUGH(intNEW As Integer) ' add go through to
    lSIZEOFlongWENTTHROUGH = lSIZEOFlongWENTTHROUGH + 1 ' add to size tracker
    If UBound(longWENTTHROUGH) < lSIZEOFlongWENTTHROUGH Then ' if not big enough to hold extra flail
        ReDim Preserve longWENTTHROUGH(0 To lSIZEOFlongWENTTHROUGH) ' resize array to hold extra flail
    End If
    longWENTTHROUGH(lSIZEOFlongWENTTHROUGH) = intNEW ' remember new monster
End Sub

Private Sub Class_Initialize() ' object init
    bACTIVE = False ' flail not on screen
    ReDim longWENTTHROUGH(0 To 10) ' resize longWENTTHROUGH to hold monsters
End Sub
