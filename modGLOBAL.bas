Attribute VB_Name = "modGLOBAL"
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

Global Const landHEIGHT = 336
Global Const numberOfMonsters = 7
Global Const ticksPerFrame = 6
Global imagePATH As String

Global arrcMONSTERPICS(0 To numberOfMonsters - 1) As New clsSPRITE
Global arrcMONSTERLPICS(0 To numberOfMonsters - 1) As New clsSPRITE

Global csprFLAIL As New clsSPRITE

Global lCURRENTLEVEL As Long

'savefile
Global strNAME As String
Global lLEVEL As Long
Global lMONEY As Long
Global intFLAILPOWER As Integer
Global lCASTLECURRENTHEALTH As Long
Global lCASTLEMAXHEALTH As Long

Public intMONSTERSONLEVEL(0 To numberOfMonsters - 1) As Integer

Public Enum monsterNames
    greenMonster
    blackMonster
    bat
    tree
    cloud
    rabbit
    ladybug
End Enum

'hDestDC - destination object, X - destination X axis, Y - destination Y axis
'nwidth - width to copy, nheight - height to copy,
'hSrcDC - source object, xSrc - start at xSrc on X axis, ySrc - start at ySrc on Y axis,
'dwRop - way to copy

Public Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal dwRop As Long _
) As Long

Public Declare Function StretchBlt Lib "gdi32.dll" ( _
    ByVal hdc As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal hSrcWidth As Long, _
    ByVal nSrcHeight As Long, _
    ByVal dwRop As Long _
) As Long

Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" ( _
    ByVal hInst As Long, _
    ByVal lpsz As String, _
    ByVal un1 As Long, _
    ByVal n1 As Long, _
    ByVal n2 As Long, _
    ByVal un2 As Long _
) As Long

Public Declare Function GetObjectW Lib "gdi32" ( _
    ByVal hObject As Long, _
    ByVal nCount As Long, _
    lpObject As Any _
) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32" ( _
    ByVal hdc As Long _
) As Long

Public Declare Function CreateCompatibleBitmap Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long _
) As Long

Public Declare Function GetDC Lib "user32" ( _
    ByVal hWnd As Long _
) As Long

Public Declare Function SelectObject Lib "gdi32.dll" ( _
    ByVal hdc As Long, _
    ByVal hObject As Long _
) As Long

Public Declare Function DeleteDC Lib "gdi32" ( _
    ByVal hdc As Long _
) As Long

Public Declare Function DeleteObject Lib "gdi32" ( _
    ByVal hObject As Long _
) As Long

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
    hpvDest As Any, _
    hpvSource As Any, _
    ByVal cbCopy As Long _
)

Public Declare Function GetBitmapBits Lib "gdi32" ( _
    ByVal hBitmap As Long, _
    ByVal dwCount As Long, _
    lpBits As Any _
) As Long

Public Declare Function SetBitmapBits Lib "gdi32" ( _
    ByVal hBitmap As Long, _
    ByVal dwCount As Long, _
    lpBits As Any _
) As Long

Public Function escapeQUOTES(strINPUT As String) As String
    escapeQUOTES = Replace$(strINPUT, "'", "''")
End Function

Public Function max(intNUM1 As Integer, intNUM2 As Integer) As Integer
    If intNUM1 < intNUM2 Then
        max = intNUM2
    Else
        max = intNUM1
    End If
End Function

Public Function min(intNUM1 As Integer, intNUM2 As Integer) As Integer
    If intNUM1 > intNUM2 Then
        min = intNUM2
    Else
        min = intNUM1
    End If
End Function

Function safeADDLONG(lNUMBER1 As Long, lNUMBER2 As Long) As Long
    Dim dblOUTPUT As Double
    dblOUTPUT = lNUMBER1
    dblOUTPUT = dblOUTPUT + lNUMBER2
    If dblOUTPUT < 2147483647 Then
        safeADDLONG = dblOUTPUT
    Else
        safeADDLONG = 2147483647
    End If
End Function

Sub main()
    ' load monster types
    Dim bSUCCESS As Boolean
    bSUCCESS = True
    
    imagePATH = App.Path & "\images\"
    
    Dim arrlIMAGESIZES(0 To numberOfMonsters - 1, 0 To 1) As Long
    
    'green monster
    arrlIMAGESIZES(0, 0) = 9
    arrlIMAGESIZES(0, 1) = 25
    arrlIMAGESIZES(1, 0) = 9
    arrlIMAGESIZES(1, 1) = 25
    arrlIMAGESIZES(2, 0) = 10
    arrlIMAGESIZES(2, 1) = 11
    arrlIMAGESIZES(3, 0) = 26
    arrlIMAGESIZES(3, 1) = 50
    arrlIMAGESIZES(4, 0) = 43
    arrlIMAGESIZES(4, 1) = 28
    arrlIMAGESIZES(5, 0) = 17
    arrlIMAGESIZES(5, 1) = 34
    arrlIMAGESIZES(6, 0) = 13
    arrlIMAGESIZES(6, 1) = 7
    
    Dim nC As Integer
    nC = 0
    Do While nC < numberOfMonsters
        bSUCCESS = bSUCCESS And arrcMONSTERPICS(nC).loadFRAMES(imagePATH & "monster" & nC & ".bmp", arrlIMAGESIZES(nC, 0), arrlIMAGESIZES(nC, 1), False, True)
        bSUCCESS = bSUCCESS And arrcMONSTERLPICS(nC).loadFRAMES(imagePATH & "monster" & nC & ".bmp", arrlIMAGESIZES(nC, 0), arrlIMAGESIZES(nC, 1), True, True)
        nC = nC + 1
    Loop
    
    ' load flail
    bSUCCESS = bSUCCESS And csprFLAIL.loadFRAMES(imagePATH & "flail.bmp", 14, 14, False, True)
    
    If bSUCCESS = False Then
        MsgBox "Error loading images!"
        End
    End If
    
    frmNEWGAME.Show
End Sub
