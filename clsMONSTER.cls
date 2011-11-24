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
Public bACTIVE As Boolean
Public sngX As Single
Public sngY As Single
Public sngMOVINGH As Single
Public intTYPE As Integer
Public intHEALTH As Integer

Private Sub Class_Initialize()
    bACTIVE = False
End Sub
