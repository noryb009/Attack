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
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

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
    
    Do While nC < UBound(bytesINBITMAP) ' for each pixel
        ' if black (rgb(0,0,0))
        If bytesINBITMAP(nC) = 0 And _
            bytesINBITMAP(nC + 1) = 0 And _
            bytesINBITMAP(nC + 2) = 0 Then
            CopyMemory bytesINBITMAP(nC), RGB(255, 255, 255), (cbitFRAMES(lFRAME).bitsPERPIXEL \ 8) ' part of background, copy white
        Else ' not black
            CopyMemory bytesINBITMAP(nC), RGB(0, 0, 0), (cbitFRAMES(lFRAME).bitsPERPIXEL \ 8) ' part of picture, copy black
        End If
        nC = nC + (cbitFRAMES(lFRAME).bitsPERPIXEL \ 8) ' next pixel
    Loop
    SetBitmapBits cbitFRAMESMASK(lFRAME).hWnd, UBound(bytesINBITMAP), bytesINBITMAP(0) ' store bytes
    makeMASK = True ' success
End Function

Public Function loadFRAMES(strFILE As String, lWIDTH As Long, lHEIGHT As Long, bFLIP As Boolean, bMASK As Boolean) As Boolean
    loadFRAMES = False ' not successful yet
    
    'TODO: remove 0, use width and height
    Dim cbitINFO As New clsBITMAP ' make a new bitmap to get info
    
    If cbitINFO.loadFILE(strFILE) = False Then Exit Function ' load image, if failure then exit
    
    If cbitINFO.width Mod lWIDTH <> 0 Or cbitINFO.height Mod lHEIGHT <> 0 Then Exit Function ' if can't get equal number of frames, exit
    
    lNUMBEROFFRAMES = (cbitINFO.width \ lWIDTH) * (cbitINFO.height \ lHEIGHT) ' get number of frames
    
    ReDim Preserve cbitFRAMES(0 To lNUMBEROFFRAMES - 1) As clsBITMAP ' make frames array big enough to hold all the frames
    ReDim Preserve cbitFRAMESMASK(0 To lNUMBEROFFRAMES - 1) As clsBITMAP ' make masks array big enough to hold all the masks
    
    Dim nC As Integer
    nC = 0
    Do While nC < lNUMBEROFFRAMES ' for each frame
        Set cbitFRAMES(nC) = New clsBITMAP ' set as new bitmap
        
        loadFRAMES = cbitFRAMES(nC).createNewImage(lWIDTH, lHEIGHT) ' create an image
        
        If loadFRAMES = False Then Exit Function ' if failure then exit
        
        BitBlt cbitFRAMES(nC).hdc, 0, 0, lWIDTH, lHEIGHT, cbitINFO.hdc, (nC * lWIDTH) Mod cbitINFO.width, ((nC * lWIDTH) \ cbitINFO.width) * lHEIGHT, vbSrcCopy ' copy frame from big image to single frame image
        
        If bFLIP = True Then ' if we should flip the image
            StretchBlt cbitFRAMES(nC).hdc, cbitFRAMES(nC).width - 1, 0, cbitFRAMES(nC).width * -1, cbitFRAMES(nC).height, cbitFRAMES(nC).hdc, 0, 0, cbitFRAMES(nC).width, cbitFRAMES(nC).height, vbSrcCopy ' flip the image
        End If
        
        If bMASK = True Then makeMASK (nC) ' if we should make a mask then make the mask
        
        If loadFRAMES = True Then ' if loaded OK
            bLOADED = True ' image has been loaded
        End If
        nC = nC + 1 ' next frame
    Loop
End Function
