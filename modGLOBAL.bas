Attribute VB_Name = "modGLOBAL"
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

Global Const SERVER = False ' not a server
Global Const programNAME = "Attack"
Global onlineMODE As Boolean ' if playing singleplayer/multiplayer

Global currentSTATE As String ' used for online mode to see what state the game is at (lobby, playing, buying, etc.)

Global cSERVER(0 To 0) As New clsCONNECTION ' connection to server
Global ccinfoPLAYERINFO(0 To MAXCLIENTS - 1) As New clsCLIENTINFO ' list of players
Global strCHATLOG(0 To 15) As String ' last few messages

Global strIMAGEPATH As String ' path to images
Global strDATABASEPATH As String ' path to data base

Global arrcMONSTERPICS(0 To numberOfMonsters - 1) As New clsSPRITE ' images of monsters going right
Global arrcMONSTERLPICS(0 To numberOfMonsters - 1) As New clsSPRITE ' images of monsters going left

Global csprFLAIL As New clsSPRITE ' flail image

Global lCURRENTLEVEL As Long ' current level
Global lHIGHSCORE As Long ' player highscore
Global lMONSTERSLEFT As Long ' only used in online mode, number of monsters left in level

'savefile
Global strNAME As String ' player name
Global lLEVEL As Long ' max level
Global lMONEY As Long ' current money
Global lLEVELMONEY As Long ' money on current level

Public intMONSTERSONLEVEL(0 To numberOfMonsters - 1) As Integer ' array with number of monsters on level

'hDestDC - destination object, X - destination X axis, Y - destination Y axis
'nwidth - width to copy, nheight - height to copy,
'hSrcDC - source object, xSrc - start at xSrc on X axis, ySrc - start at ySrc on Y axis,
'dwRop - way to copy
' copy part of an image to another
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

' copy part of an image to another and strech
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

' load image from a file
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" ( _
    ByVal hInst As Long, _
    ByVal lpsz As String, _
    ByVal un1 As Long, _
    ByVal n1 As Long, _
    ByVal n2 As Long, _
    ByVal un2 As Long _
) As Long

' get hWnd handle of object
Public Declare Function GetObjectW Lib "gdi32" ( _
    ByVal hObject As Long, _
    ByVal nCount As Long, _
    lpObject As Any _
) As Long

' create a hDC
Public Declare Function CreateCompatibleDC Lib "gdi32" ( _
    ByVal hdc As Long _
) As Long

' create a bitmap
Public Declare Function CreateCompatibleBitmap Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long _
) As Long

' get the hDC of an hWnd
Public Declare Function GetDC Lib "user32" ( _
    ByVal hWnd As Long _
) As Long

' bind an hDC and an object together
Public Declare Function SelectObject Lib "gdi32.dll" ( _
    ByVal hdc As Long, _
    ByVal hObject As Long _
) As Long

' delete an hDC
Public Declare Function DeleteDC Lib "gdi32" ( _
    ByVal hdc As Long _
) As Long

' delete an hWnd
Public Declare Function DeleteObject Lib "gdi32" ( _
    ByVal hObject As Long _
) As Long

' copy bytes from one location to another
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
    hpvDest As Any, _
    hpvSource As Any, _
    ByVal cbCopy As Long _
)

' get the image part of a bitmap
Public Declare Function GetBitmapBits Lib "gdi32" ( _
    ByVal hBitmap As Long, _
    ByVal dwCount As Long, _
    lpBits As Any _
) As Long

' set the image part of a bitmap
Public Declare Function SetBitmapBits Lib "gdi32" ( _
    ByVal hBitmap As Long, _
    ByVal dwCount As Long, _
    lpBits As Any _
) As Long

' get the current number of ticks
Public Declare Function QueryPerformanceCounter Lib "kernel32" ( _
    lpPerformanceCount As Currency _
) As Long

' get the frequency of ticks
Public Declare Function QueryPerformanceFrequency Lib "kernel32" ( _
    lpFrequency As Currency _
) As Long

' sleep, don't use CPU
Public Declare Sub Sleep Lib "kernel32" ( _
    ByVal dwMilliseconds As Long _
)

' escape ' to use in SQL queries
Public Function escapeQUOTES(strINPUT As String) As String
    escapeQUOTES = Replace$(strINPUT, "'", "''")
End Function

' load monster info into cmontypeMONSTERINFO
Sub loadONEMONSTERINFO(intNUMBER As Integer, imageNAME As String, lIMAGEWIDTH As Long, lIMAGEHEIGHT As Long, intPOINTCOST As Integer, intHEALTH As Integer, intATTACKPOWER As Integer, intSTARTINGY As Integer, sngXSPEED As Single, intMONEYONHIT As Integer, intMONEYONKILL As Integer, Optional sngYSPEED As Single = 0)
    Dim bSUCCESS As Boolean ' successful
    bSUCCESS = True ' default: true
    
    bSUCCESS = bSUCCESS And arrcMONSTERPICS(intNUMBER).loadFRAMES(strIMAGEPATH & imageNAME & ".bmp", lIMAGEWIDTH, lIMAGEHEIGHT, False, True) ' load image
    bSUCCESS = bSUCCESS And arrcMONSTERLPICS(intNUMBER).loadFRAMES(strIMAGEPATH & imageNAME & ".bmp", lIMAGEWIDTH, lIMAGEHEIGHT, True, True) ' load image looking left
    
    If bSUCCESS = False Then ' if error
        MsgBox "Error loading images!", vbOKOnly, programNAME ' alert user
        End ' exit program
    End If
    
    cmontypeMONSTERINFO(intNUMBER).intPOINTCOST = intPOINTCOST ' load point cost
    cmontypeMONSTERINFO(intNUMBER).intMAXHEALTH = intHEALTH ' load health
    cmontypeMONSTERINFO(intNUMBER).intATTACKPOWER = intATTACKPOWER ' load attack power
    cmontypeMONSTERINFO(intNUMBER).sngXSPEED = sngXSPEED ' load vertical speed
    cmontypeMONSTERINFO(intNUMBER).sngYSPEED = sngYSPEED ' load horizontal speed
    If intSTARTINGY = -1 Then ' default: ground
        cmontypeMONSTERINFO(intNUMBER).intSTARTINGY = landHEIGHT - arrcMONSTERPICS(intNUMBER).height ' Y is land height - image height, so feet are on ground
    Else ' special Y location
        cmontypeMONSTERINFO(intNUMBER).intSTARTINGY = intSTARTINGY ' load Y location
    End If
    cmontypeMONSTERINFO(intNUMBER).intMONEYADDEDHIT = intMONEYONHIT ' load money on hit
    cmontypeMONSTERINFO(intNUMBER).intMONEYADDEDKILL = intMONEYONKILL ' load money on kill
    cmontypeMONSTERINFO(intNUMBER).lWIDTH = lIMAGEWIDTH ' load image width
    cmontypeMONSTERINFO(intNUMBER).lHEIGHT = lIMAGEHEIGHT ' load image height
    cmontypeMONSTERINFO(intNUMBER).lFRAMES = arrcMONSTERPICS(intNUMBER).numberOfFrames ' load number of frames
End Sub

Sub Main()
    Randomize ' randomize random numbers
    
    strIMAGEPATH = App.Path & "\images\" ' images are found in the images folder
    strDATABASEPATH = App.Path & "\saveFiles.mdb" ' database is in the same folder as the EXE
    
    ' load monster types
    Dim bSUCCESS As Boolean
    bSUCCESS = True ' successful so far
    
    loadMONSTERINFO ' load monster info into cmontypeMONSTERINFO()
    loadPLAYERCOLOURS ' load player colour info into playerCOLOURS()
    
    bSUCCESS = bSUCCESS And csprFLAIL.loadFRAMES(strIMAGEPATH & "flail.bmp", 14, 14, False, True) ' load flail image
    If csprFLAIL.numberOfFrames < MAXCLIENTS + 1 Then ' if not enough frames
        bSUCCESS = False ' error
    End If
    
    If bSUCCESS = False Then ' if error
        MsgBox "Error loading images!", vbOKOnly, programNAME ' alert user
        End ' exit program
    End If
    
    frmNEWGAME.Show ' show new game form
End Sub

Public Sub broadcast(strCOMMAND As String, strTOSEND As String) ' send a command to all the clients
     ' empty sub to allow server to share subs
End Sub
Public Sub log(strNEWLINE As String)
    ' empty sub to allow server to share subs
End Sub
Public Sub broadcastMONSTER(lMONSTERNUMBER As Long)
    ' empty sub to allow server to share subs
End Sub
Public Sub broadcastFLAIL(lFLAILNUMBER As Long, bCLEARGOTHROUGH As Boolean)
    ' empty sub to allow server to share subs
End Sub
