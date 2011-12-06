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
Dim csprMONSTER As clsSPRITE

Public strNAME As String
Public sngY As Single
Public sngMOVINGH As Single
Public intHEALTH As Integer
Public lTOTALFRAMES As Long
Public lANIMATIONRATE As Long

Public Function openSPRITE(strFILENAME As String) As Boolean
    openSPRITE = True 'clssprite.loadFRAMES strFILENAME
End Function
