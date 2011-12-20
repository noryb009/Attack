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
Public intGOTHROUGH As Integer
Dim longWENTTHROUGH() As Integer

Public Sub clearWENTTHROUGH()
    ReDim longWENTTHROUGH(0 To 0)
End Sub

Public Function didGOTHROUGH(intCHECK As Integer) As Boolean
    Dim nC As Integer
    nC = 1
    didGOTHROUGH = False
    Do While nC <= UBound(longWENTTHROUGH)
        If longWENTTHROUGH(nC) = intCHECK Then
            didGOTHROUGH = True
            Exit Do
        End If
        nC = nC + 1
    Loop
End Function

Public Sub addGOTHROUGH(intNEW As Integer)
    ReDim Preserve longWENTTHROUGH(0 To UBound(longWENTTHROUGH) + 1)
    longWENTTHROUGH(UBound(longWENTTHROUGH)) = intNEW
End Sub

Private Sub Class_Initialize()
    bACTIVE = False
    ReDim longWENTTHROUGH(0 To 0)
End Sub
