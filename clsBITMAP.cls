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

Const LR_LOADFROMFILE = &H10 ' needed for LoadImage

Dim lWND As Long ' WND handle - the handle LoadImage returns
Dim lHDC As Long ' hDC handle - the handle needed for BitBlt
Dim bitmapINFO As bitmapSTRUCT ' the info about the bitmap

Private Type bitmapSTRUCT ' all the parts of the bitmap header
    bmTYPE As Long ' always 0 if loaded correctly
    bmWIDTH As Long ' image width
    bmHEIGHT As Long ' image height
    bmWIDTHBYTES As Long ' bytes in one row
    bmPLANES As Integer ' planes in the image
    bmBITSPIXEL As Integer ' bits per pixel
    bmBITS As Long ' total size of bitmap in bits
End Type

Private Sub Class_Terminate()
    If lHDC <> 0 Then ' if hDC is set
        DeleteDC lHDC ' remove the hDC handle
        lHDC = 0 ' set to 0
    End If
    If lWND <> 0 Then ' if WND is set
        DeleteObject lWND ' remove the WND handle
        lWND = 0 ' set to 0
    End If
End Sub

Public Property Get bits()
    bits = bitmapINFO.bmBITS ' return the bits in the image
End Property

Public Property Get bitsPERPIXEL()
    bitsPERPIXEL = bitmapINFO.bmBITSPIXEL ' return the bits per pixel
End Property

Public Property Get hdc()
    hdc = lHDC ' return the hDC
End Property

Public Property Get height()
    height = bitmapINFO.bmHEIGHT ' return the height
End Property

Public Property Get hWnd()
    hWnd = lWND ' return the WND
End Property

Public Property Get width()
    width = bitmapINFO.bmWIDTH ' return the width
End Property

Public Property Get widthBYTES()
    widthBYTES = bitmapINFO.bmWIDTHBYTES ' return the bits in one row
End Property

Public Function loadFILE(strFILENAME As String) As Boolean
    loadFILE = False ' didn't load successfully
    
    On Error GoTo endOFloadFILE
    
    ' if object already exists, delete it
    If lHDC <> 0 Then
        DeleteDC lHDC ' delete old hDC
        lHDC = 0
    End If
    If lWND <> 0 Then
        DeleteObject lWND ' delete old lWND
        lWND = 0
    End If
    
    lWND = LoadImage(0, strFILENAME, 0, 0, 0, LR_LOADFROMFILE) ' load the image
    
    GetObjectW lWND, Len(bitmapINFO), bitmapINFO ' get the image info and put it in bitmapINFO
    If lWND = 0 Or bitmapINFO.bmBITSPIXEL = 0 Or bitmapINFO.bmTYPE <> 0 Then ' if didn't load successfully
        Exit Function
    End If
    
    lHDC = CreateCompatibleDC(0) ' make an hDC
    If lHDC = 0 Then ' if wasn't created
        Exit Function
    End If
    
    SelectObject lHDC, lWND ' set the hDC to point to the WND
    
    loadFILE = True ' loaded successfully
    Exit Function
endOFloadFILE: ' on error
    loadFILE = False ' didn't load successfully
End Function

Public Function createNewImage(width As Long, height As Long) As Boolean
    createNewImage = False ' didn't load successfully
    
    On Error GoTo endOFcreateNewImage
    
    ' if object already exists, delete it
    If lHDC <> 0 Then
        DeleteDC lHDC ' delete old hDC
        lHDC = 0
    End If
    If lWND <> 0 Then
        DeleteObject lWND ' delete old lWND
        lWND = 0
    End If
    
    lHDC = CreateCompatibleDC(GetDC(0)) ' make new hDC
    If lHDC = 0 Then ' if empty
        Exit Function
    End If
    
    lWND = CreateCompatibleBitmap(GetDC(0), width, height) ' make bitmap and point WND to it
    If lWND = 0 Then ' if wasn't loaded correctly
        Exit Function
    End If
    
    SelectObject lHDC, lWND ' set the hDC to point to the WND
    
    GetObjectW lWND, Len(bitmapINFO), bitmapINFO ' get the image info and put it in bitmapINFO
    
    createNewImage = True ' loaded successfully
    Exit Function
endOFcreateNewImage: ' on error
    createNewImage = False ' didn't create image, return error
End Function
