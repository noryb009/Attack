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
Dim bLOADED As Boolean ' is loaded
Dim cbitFRAMES() As clsBITMAP ' array of frames in sprite
Dim cbitFRAMESMASK() As clsBITMAP ' array of frame masks in sprite
Dim lNUMBEROFFRAMES As Long ' number of frames in animation

Public Property Get framehDC(intFRAME As Integer)
    framehDC = cbitFRAMES(intFRAME).hdc ' return hDC handle
End Property

Public Property Get frameMaskhDC(intFRAME As Integer)
    frameMaskhDC = cbitFRAMESMASK(intFRAME).hdc ' return mask hDC handle
End Property

Public Property Get width()
    width = cbitFRAMES(0).width ' return frame widths
End Property

Public Property Get height()
    height = cbitFRAMES(0).height ' return frame heights
End Property

Public Property Get numberOfFrames()
    numberOfFrames = lNUMBEROFFRAMES ' return number of frames
End Property

Private Sub Class_Initialize()
    bLOADED = False ' default: image not loaded
End Sub

Private Function makeMASK(lFRAME As Long) As Boolean
    makeMASK = False ' didn't make the mask successfully
    Set cbitFRAMESMASK(lFRAME) = New clsBITMAP ' create a new bitmap
    If cbitFRAMESMASK(lFRAME).createNewImage(cbitFRAMES(lFRAME).width, cbitFRAMES(lFRAME).height) = False Then ' make an empty image
        Exit Function
    End If
    BitBlt cbitFRAMESMASK(lFRAME).hdc, 0, 0, cbitFRAMES(lFRAME).width, cbitFRAMES(lFRAME).height, cbitFRAMES(lFRAME).hdc, 0, 0, vbSrcCopy ' copy frame into mask
    
    Dim bytesINBITMAP() As Byte ' array for bytes in bitmap
    ReDim bytesINBITMAP(0 To cbitFRAMES(lFRAME).widthBYTES * cbitFRAMES(lFRAME).height - 1) ' resize the array to hole the bytes
    
    GetBitmapBits cbitFRAMESMASK(lFRAME).hWnd, UBound(bytesINBITMAP), bytesINBITMAP(0) ' get the bytes into the array, starting at bytesINBITMAP(0)
    
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
    Dim cbitINFO As New clsBITMAP
    
    If cbitINFO.loadFILE(strFILE) = False Then Exit Function
    
    If cbitINFO.width Mod lWIDTH <> 0 Or cbitINFO.height Mod lHEIGHT <> 0 Then Exit Function ' if can't get equal number of frames, exit
    
    lNUMBEROFFRAMES = (cbitINFO.width \ lWIDTH) * (cbitINFO.height \ lHEIGHT)
    
    ReDim Preserve cbitFRAMES(0 To lNUMBEROFFRAMES - 1) As clsBITMAP
    ReDim Preserve cbitFRAMESMASK(0 To lNUMBEROFFRAMES - 1) As clsBITMAP
    Dim nC As Integer
    nC = 0
    Do While nC < lNUMBEROFFRAMES
        Set cbitFRAMES(nC) = New clsBITMAP
        'loadFRAMES = cbitFRAMES(nC).loadFILE(strFILE)
        
        loadFRAMES = cbitFRAMES(nC).createNewImage(lWIDTH, lHEIGHT)
        
        If loadFRAMES = False Then Exit Function
        
        BitBlt cbitFRAMES(nC).hdc, 0, 0, lWIDTH, lHEIGHT, cbitINFO.hdc, (nC * lWIDTH) Mod cbitINFO.width, nC \ cbitINFO.width, vbSrcCopy
        
        If bFLIP = True Then
            StretchBlt cbitFRAMES(nC).hdc, cbitFRAMES(nC).width - 1, 0, cbitFRAMES(nC).width * -1, cbitFRAMES(nC).height, cbitFRAMES(nC).hdc, 0, 0, cbitFRAMES(nC).width, cbitFRAMES(nC).height, vbSrcCopy
        End If
        
        If bMASK = True Then makeMASK (nC)
        
        If loadFRAMES = True Then
            bLOADED = True
        End If
        nC = nC + 1
    Loop
End Function
