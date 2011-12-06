VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSPRITE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim bLOADED As Boolean
Dim cbitFRAMES() As clsBITMAP
Dim cbitFRAMESMASK() As clsBITMAP
Dim intWIDTH As Integer
Dim intHEIGHT As Integer
Dim lNUMBEROFFRAMES As Long

Public Property Get framehDC(intFRAME As Integer)
    framehDC = cbitFRAMES(intFRAME).hdc
End Property

Public Property Get frameMaskhDC(intFRAME As Integer)
    frameMaskhDC = cbitFRAMESMASK(intFRAME).hdc
End Property

Public Property Get width()
    width = cbitFRAMES(0).width
End Property

Public Property Get height()
    height = cbitFRAMES(0).height
End Property

Public Property Get numberOfFrames()
    numberOfFrames = lNUMBEROFFRAMES
End Property

Private Sub Class_Initialize()
    bLOADED = False
End Sub

Private Function makeMASK(lFRAME As Long) As Boolean
    makeMASK = False
    ReDim Preserve cbitFRAMESMASK(0 To lFRAME) As clsBITMAP
    Set cbitFRAMESMASK(lFRAME) = New clsBITMAP
    If cbitFRAMESMASK(lFRAME).createNewImage(cbitFRAMES(lFRAME).width, cbitFRAMES(lFRAME).height) = False Then
        Exit Function
    End If
    BitBlt cbitFRAMESMASK(lFRAME).hdc, 0, 0, cbitFRAMES(lFRAME).width, cbitFRAMES(lFRAME).height, cbitFRAMES(lFRAME).hdc, 0, 0, vbSrcCopy
    
    Dim bytesINBITMAP() As Byte
    ReDim bytesINBITMAP(0 To cbitFRAMES(lFRAME).widthBYTES * cbitFRAMES(lFRAME).height - 1)
    
    GetBitmapBits cbitFRAMESMASK(lFRAME).hWnd, UBound(bytesINBITMAP), bytesINBITMAP(0)
    
    Dim nC As Long
    nC = 0
    
    Do While nC < UBound(bytesINBITMAP)
        If bytesINBITMAP(nC) = 0 And _
            bytesINBITMAP(nC + 1) = 0 And _
            bytesINBITMAP(nC + 2) = 0 Then
            CopyMemory bytesINBITMAP(nC), RGB(255, 255, 255), (cbitFRAMES(lFRAME).bitsPERPIXEL \ 8)
        Else
            CopyMemory bytesINBITMAP(nC), RGB(0, 0, 0), (cbitFRAMES(lFRAME).bitsPERPIXEL \ 8)
        End If
        nC = nC + (cbitFRAMES(lFRAME).bitsPERPIXEL \ 8)
    Loop
    SetBitmapBits cbitFRAMESMASK(lFRAME).hWnd, UBound(bytesINBITMAP), bytesINBITMAP(0)
    makeMASK = True
End Function

Public Function loadFRAMES(strFILE As String, lWIDTH As Long, lHEIGHT As Long, bFLIP As Boolean, bMASK As Boolean) As Boolean
    loadFRAMES = False
    
    'TODO: remove 0, use width and height
    lNUMBEROFFRAMES = 1
    ReDim Preserve cbitFRAMES(0 To 0) As clsBITMAP
    Set cbitFRAMES(0) = New clsBITMAP
    loadFRAMES = cbitFRAMES(0).loadFILE(strFILE)
    If loadFRAMES = False Then Exit Function
    If bFLIP = True Then
        StretchBlt cbitFRAMES(0).hdc, cbitFRAMES(0).width, 0, cbitFRAMES(0).width * -1, cbitFRAMES(0).height, cbitFRAMES(0).hdc, 0, 0, cbitFRAMES(0).width, cbitFRAMES(0).height, vbSrcCopy
    End If
    
    If bMASK = True Then makeMASK (0)

    If loadFRAMES = True Then
        bLOADED = True
    End If
End Function
