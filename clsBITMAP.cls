VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBITMAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

Const LR_LOADFROMFILE = &H10

Dim lWND As Long
Dim lHDC As Long
Dim bitmapINFO As bitmapSTRUCT

Private Type bitmapSTRUCT
    bmTYPE As Long
    bmWIDTH As Long
    bmHEIGHT As Long
    bmWIDTHBYTES As Long
    bmPLANES As Integer
    bmBITSPIXEL As Integer
    bmBITS As Long
End Type

Private Sub Class_Terminate()
    If lHDC <> 0 Then
        DeleteDC lHDC
    End If
    If lWND <> 0 Then
        DeleteObject lWND
    End If
End Sub

Public Property Get bits()
    bits = bitmapINFO.bmBITS
End Property

Public Property Get bitsPERPIXEL()
    bitsPERPIXEL = bitmapINFO.bmBITSPIXEL
End Property

Public Property Get hdc()
    hdc = lHDC
End Property

Public Property Get height()
    height = bitmapINFO.bmHEIGHT
End Property

Public Property Get hWnd()
    hWnd = lWND
End Property

Public Property Get width()
    width = bitmapINFO.bmWIDTH
End Property

Public Property Get widthBYTES()
    widthBYTES = bitmapINFO.bmWIDTHBYTES
End Property

Public Function loadFILE(strFILENAME As String) As Boolean
    loadFILE = False
    
    ' if object already exists, delete it
    If lHDC <> 0 Then DeleteDC lHDC
    If lWND <> 0 Then DeleteObject lWND
    
    'On Error GoTo endOFloadFILE TODO: uncomment
    lWND = LoadImage(0, strFILENAME, 0, 0, 0, LR_LOADFROMFILE)
    GetObjectW lWND, Len(bitmapINFO), bitmapINFO
    If lWND = 0 Or bitmapINFO.bmBITSPIXEL = 0 Or bitmapINFO.bmTYPE <> 0 Then
        Exit Function
    End If
    
    lHDC = CreateCompatibleDC(0)
    If lHDC = 0 Then
        Exit Function
    End If
    
    SelectObject lHDC, lWND
    
    loadFILE = True
    Exit Function
endOFloadFILE:
    MsgBox Err.Number
End Function

Public Function createNewImage(width As Long, height As Long) As Boolean
    createNewImage = False
    
    ' if object already exists, delete it
    If lWND <> 0 Then DeleteObject lWND
    If lHDC <> 0 Then DeleteDC lHDC
    
    lHDC = CreateCompatibleDC(GetDC(0))
    If lHDC = 0 Then
        Exit Function
    End If
    
    lWND = CreateCompatibleBitmap(GetDC(0), width, height)
    If lWND = 0 Then
        Exit Function
    End If
    
    SelectObject lHDC, lWND
    
    GetObjectW lWND, Len(bitmapINFO), bitmapINFO
    
    createNewImage = True
End Function
