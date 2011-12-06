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

Public bACTIVE As Boolean
Public sngX As Single
Public sngY As Single
Public sngMOVINGH As Single
Public sngMOVINGV As Single
Public lCURRENTANIFRAME As Long

Public Sub nextFRAME()
    lCURRENTANIFRAME = lCURRENTANIFRAME + 1
    If lCURRENTANIFRAME = csprFLAIL.numberOfFrames Then
        lCURRENTANIFRAME = 0
    End If
End Sub

Private Sub Class_Initialize()
    bACTIVE = False
End Sub
