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

Public bACTIVE As Boolean
Public sngX As Single
Public sngY As Single
Public sngMOVINGH As Single
Public intTYPE As Integer
Public intHEALTH As Integer
Dim lCURRENTANIFRAME As Long

Public Property Get currentFRAME() As Long
    currentFRAME = lCURRENTANIFRAME \ ticksPerFrame
End Property

Public Property Let currentFRAME(ByVal newFRAME As Long)
    If newFRAME > -1 And newFRAME < arrcMONSTERPICS(intTYPE).numberOfFrames * ticksPerFrame Then
        lCURRENTANIFRAME = newFRAME \ ticksPerFrame
    Else
        lCURRENTANIFRAME = 0
    End If
End Property

Public Property Let fromTYPE(ByVal intNEWTYPE As Integer)
    bACTIVE = True
    intTYPE = intNEWTYPE
End Property

Private Sub Class_Initialize()
    bACTIVE = False
End Sub

Public Sub nextFRAME()
    lCURRENTANIFRAME = lCURRENTANIFRAME + 1
    If lCURRENTANIFRAME = arrcMONSTERPICS(intTYPE).numberOfFrames * ticksPerFrame Then
        lCURRENTANIFRAME = 0
    End If
End Sub

Public Sub previousFRAME()
    lCURRENTANIFRAME = lCURRENTANIFRAME - 1
    If lCURRENTANIFRAME = -1 Then
        lCURRENTANIFRAME = (arrcMONSTERPICS(intTYPE).numberOfFrames * ticksPerFrame) - 1
    End If
End Sub

