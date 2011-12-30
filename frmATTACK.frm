VERSION 5.00
Begin VB.Form frmATTACK 
   AutoRedraw      =   -1  'True
   Caption         =   "Attack"
   ClientHeight    =   1380
   ClientLeft      =   825
   ClientTop       =   1365
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   ScaleHeight     =   92
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   278
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timerMAIN 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   960
      Top             =   600
   End
   Begin VB.Label lblSCORE 
      BackStyle       =   0  'Transparent
      Caption         =   "lblSCORE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Line lineAIM 
      Visible         =   0   'False
      X1              =   8
      X2              =   88
      Y1              =   32
      Y2              =   32
   End
End
Attribute VB_Name = "frmATTACK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

' images
Dim cbitBACKGROUND As New clsBITMAP ' static background
Dim csprCASTLE As New clsSPRITE ' castle with different health ranges
Dim cbitBUFFER As New clsBITMAP ' buffer
Dim cbitHEALTH As New clsBITMAP ' health bar

Const keepX = 338 ' X location of flail starting point
Const keepY = 190 ' Y location of flail starting point

Const castleTOPMARGIN = 150 ' space above top of castle image

Private Sub Form_Activate()
    Dim currSTARTTIME As Currency ' starting time
    Dim currCURRENTTIME As Currency ' current time
    'Dim currFREQUENCY As Currency ' frame frequency
    Dim dblTIMEBETWEENFRAMES As Double ' time between frames
    
    QueryPerformanceFrequency currCURRENTTIME ' currFREQUENCY ' get the frequency of ticks
    dblTIMEBETWEENFRAMES = currCURRENTTIME / FPS ' currFREQUENCY / FPS ' get time between frames needed to reach FPS
    
    Dim bDRAWN As Boolean
    bDRAWN = False
    
    Do While bEXIT = False And bFORCEEXIT = False
        QueryPerformanceCounter currCURRENTTIME ' get current time
        If currCURRENTTIME >= currSTARTTIME + dblTIMEBETWEENFRAMES Then ' if start time + time between frame = current time, then time for the next frame
            QueryPerformanceCounter currSTARTTIME ' store current time as new start time
            moveEVERYTHING ' move everything
            'drawEVERYTHING ' draw everything
            If onlineMODE = False Then checkWINLOOSE ' run win/loose code
            bDRAWN = False ' you haven't drawn this frame yet
        Else
            If bDRAWN = True Then ' if frame already drawn
                Sleep 1 ' sleep
            Else
                drawEVERYTHING ' draw everything
                bDRAWN = True ' you have drawn this frame
            End If
        End If
        DoEvents ' do any events needed to be done
    Loop
    
    drawEVERYTHING ' draw everything a final time
    
    Unload frmATTACK ' hide this frame
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then ' if left click
        lineAIM.Visible = True ' show the aim line
        lineAIM.X1 = x ' set the starting x
        lineAIM.Y1 = y ' set the starting y
        lineAIM.X2 = x ' set the current x
        lineAIM.Y2 = y ' set the current y
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Or Button = 3 Or Button = 5 Or Button = 7 Then
        lineAIM.X2 = x
        lineAIM.Y2 = y
    End If
End Sub

Private Sub Form_Resize()
'If frmATTACK.WindowState <> vbMinimized Then ' if minimized, don't worry
'    If frmATTACK.WindowState = vbMaximized Then ' if maximized
'        frmATTACK.WindowState = vbNormal ' return to window mode
'    End If
'    frmATTACK.width = (windowX + (frmATTACK.width / Screen.TwipsPerPixelX) - frmATTACK.ScaleWidth) * Screen.TwipsPerPixelX ' width = width + border
'    frmATTACK.height = (windowY + (frmATTACK.height / Screen.TwipsPerPixelY) - frmATTACK.ScaleHeight) * Screen.TwipsPerPixelY ' height = height + border
'End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Const divideSPEED = 10

If Button = 1 Then ' mouse released
    lineAIM.Visible = False ' make aim line invisible
    
    If (lineAIM.X1 - lineAIM.X2) \ divideSPEED = 0 And (lineAIM.Y1 - lineAIM.Y2) \ divideSPEED = 0 Then
        Exit Sub
    End If
    
    Dim nC As Integer
    Dim nC2 As Integer
    Dim nCMAX As Integer
    nC = 0
    nC2 = 0
    Do While nC2 < intFLAILAMOUNT
        If onlineMODE = False Then
            nCMAX = UBound(arrFLAILS)
            Do While nC <= nCMAX
                If arrFLAILS(nC).bACTIVE = False Then
                    arrFLAILS(nC).bACTIVE = True
                    arrFLAILS(nC).sngX = keepX
                    arrFLAILS(nC).sngY = keepY
                    arrFLAILS(nC).sngMOVINGV = (lineAIM.Y1 - lineAIM.Y2) \ divideSPEED
                    arrFLAILS(nC).sngMOVINGH = (lineAIM.X1 - lineAIM.X2) \ divideSPEED
                    
                    arrFLAILS(nC).sngMOVINGV = arrFLAILS(nC).sngMOVINGV + (((intFLAILAMOUNT / 2) - 0.5 - nC2) * 4)
                    'arrFLAILS(nC).sngMOVINGH = arrFLAILS(nC).sngMOVINGH + (((intFLAILAMOUNT / 2) - 0.5 - nC2) * 2)
                    
                    arrFLAILS(nC).lCURRENTANIFRAME = 0
                    arrFLAILS(nC).intGOTHROUGH = intFLAILGOTHROUGH
                    arrFLAILS(nC).clearWENTTHROUGH
                    Exit Do
                End If
                nC = nC + 1
            Loop
        Else
            cSERVER(0).sendString "newFla", True & "~" & keepX & "~" & keepY & "~" & _
            (lineAIM.Y1 - lineAIM.Y2) \ divideSPEED + (((intFLAILAMOUNT / 2) - 0.5 - nC2) * 4) & "~" & _
            (lineAIM.X1 - lineAIM.X2) \ divideSPEED & "~" & _
            intFLAILGOTHROUGH & "~" & _
            True
        End If
        nC2 = nC2 + 1
    Loop
    Else
End If
End Sub

Sub drawBUFFER()
frmATTACK.Cls

'BitBlt frmATTACK.hdc, 0, 0, cbitBUFFER.width, cbitBUFFER.height, cbitBUFFER.hdc, 0, 0, vbSrcCopy
StretchBlt frmATTACK.hdc, 0, 0, frmATTACK.ScaleWidth, frmATTACK.ScaleHeight, cbitBUFFER.hdc, 0, 0, cbitBUFFER.width, cbitBUFFER.height, vbSrcCopy

frmATTACK.Refresh
End Sub

Sub spawnMONSTER()
    Dim nCMAX As Integer
    If intCURRENTMONSTER <= UBound(arrTOBEMONSTERS) Then
        nC = 0
        nCMAX = UBound(arrMONSTERS)
        Do While nC <= nCMAX
            If arrMONSTERS(nC).bACTIVE = False Then
                arrMONSTERS(nC).bACTIVE = True
                arrMONSTERS(nC).intTYPE = arrTOBEMONSTERS(intCURRENTMONSTER) 'Int(Rnd() * numberOfMonsters)
                arrMONSTERS(nC).currentFRAME = 0
                
                arrMONSTERS(nC).sngX = Int(Rnd() * 2)
                If arrMONSTERS(nC).sngX = 0 Then
                    arrMONSTERS(nC).sngX = 0 - arrcMONSTERPICS(arrMONSTERS(nC).intTYPE).width
                    arrMONSTERS(nC).sngMOVINGH = 1 * cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).sngSPEED ' go left
                Else
                    arrMONSTERS(nC).sngX = windowX + arrcMONSTERPICS(arrMONSTERS(nC).intTYPE).width
                    arrMONSTERS(nC).sngMOVINGH = -1 * cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).sngSPEED ' go left
                End If
                arrMONSTERS(nC).sngY = cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).intSTARTINGY
                arrMONSTERS(nC).intHEALTH = cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).intMAXHEALTH
                
                Exit Do
            End If
            nC = nC + 1
        Loop
        intCURRENTMONSTER = intCURRENTMONSTER + 1
    End If
End Sub

Sub moveEVERYTHING()
    Dim nC As Long
    Dim nCMAXMON As Long
    nCMAXMON = UBound(arrMONSTERS)
    Dim nCMAXFLAILS As Long
    nCMAXFLAILS = UBound(arrFLAILS)
    
    ' spawn monsters
    If onlineMODE = False Then
        If lMONSTERSPAWNCOOLDOWN = 0 Then
            Dim bSPAWN As Boolean
            bSPAWN = False
            If intCURRENTMONSTER <= intMONSTERSKILLED + intMONSTERSATTACKEDCASTLE + (lCURRENTLEVEL \ 3) Then ' force if nobody on screen
                bSPAWN = True
            ElseIf Int(Rnd() * 200) < lCURRENTLEVEL And intCURRENTMONSTER <= UBound(arrTOBEMONSTERS) Then ' randomly if some monsters are waiting
                bSPAWN = True
            End If
            
            If bSPAWN = True Then
                spawnMONSTER
                lMONSTERSPAWNCOOLDOWN = 20
            End If
        Else
            lMONSTERSPAWNCOOLDOWN = lMONSTERSPAWNCOOLDOWN - 1
        End If
    End If
    
    ' move monsters
    nC = 0
    Do While nC <= nCMAXMON
        If arrMONSTERS(nC).bACTIVE = True Then
            arrMONSTERS(nC).moveMONSTER
            If (arrMONSTERS(nC).sngMOVINGH < 0 And arrMONSTERS(nC).sngX + arrcMONSTERPICS(arrMONSTERS(nC).intTYPE).width < 0) Or (arrMONSTERS(nC).sngMOVINGH > 0 And arrMONSTERS(nC).sngX > windowX) Then
                arrMONSTERS(nC).bACTIVE = False
            ElseIf (arrMONSTERS(nC).sngMOVINGH > 0 And arrMONSTERS(nC).sngX + arrcMONSTERPICS(arrMONSTERS(nC).intTYPE).width > castleWALLLEFT) Or (arrMONSTERS(nC).sngMOVINGH < 0 And arrMONSTERS(nC).sngX < castleWALLRIGHT) Then 'attack
                lCASTLECURRENTHEALTH = lCASTLECURRENTHEALTH - cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).intATTACKPOWER
                intMONSTERSATTACKEDCASTLE = intMONSTERSATTACKEDCASTLE + 1
                arrMONSTERS(nC).bACTIVE = False
                If lCASTLECURRENTHEALTH <= 0 And onlineMODE = False Then bEXIT = True
            End If
        End If
        nC = nC + 1
    Loop
    
    If lCASTLECURRENTHEALTH <= 0 And onlineMODE = False Then ' if dead and not online
        bEXIT = True ' exit
    End If
    
    ' move flails
    nC = 0
    Do While nC <= nCMAXFLAILS
        If arrFLAILS(nC).bACTIVE = True Then
            Dim intNEWX As Integer
            Dim intNEWY As Integer
            Dim intDELETEMONSTER As Integer
            intDELETEMONSTER = -1 ' don't delete a monster
            intNEWX = Int(arrFLAILS(nC).sngX) + arrFLAILS(nC).sngMOVINGH
            intNEWY = Int(arrFLAILS(nC).sngY) + arrFLAILS(nC).sngMOVINGV
            
            'gravity
            arrFLAILS(nC).sngMOVINGV = arrFLAILS(nC).sngMOVINGV + 0.5
            If arrFLAILS(nC).sngMOVINGH < 0 Then
                arrFLAILS(nC).sngMOVINGH = arrFLAILS(nC).sngMOVINGH + 0.1
            ElseIf arrFLAILS(nC).sngMOVINGH > 0 Then
                arrFLAILS(nC).sngMOVINGH = arrFLAILS(nC).sngMOVINGH - 0.1
            End If
            
            'check if hitting monster
            Dim nCMONSTERS As Integer
            nCMONSTERS = 0
            Do While nCMONSTERS <= nCMAXMON ' UBound(arrMONSTERS)
                If arrMONSTERS(nCMONSTERS).bACTIVE = True Then
                    If (arrFLAILS(nC).sngY < arrMONSTERS(nCMONSTERS).sngY + arrcMONSTERPICS(arrMONSTERS(nC).intTYPE).height And intNEWY + arrcMONSTERPICS(arrMONSTERS(nC).intTYPE).height > arrMONSTERS(nCMONSTERS).sngY) Or _
                    (intNEWY < arrMONSTERS(nCMONSTERS).sngY + arrcMONSTERPICS(arrMONSTERS(nC).intTYPE).height And arrFLAILS(nC).sngY + csprFLAIL.height > arrMONSTERS(nCMONSTERS).sngY) Then
                        'If arrFLAILS(nC).sngx + picFLAIL.Width < arrMONSTERS(nCMONSTERS).sngx And intNEWX + picFLAIL.Width > arrMONSTERS(nCMONSTERS).sngx Then
                        If (arrFLAILS(nC).sngX < arrMONSTERS(nCMONSTERS).sngX + arrcMONSTERPICS(arrMONSTERS(nC).intTYPE).width And intNEWX + csprFLAIL.width > arrMONSTERS(nCMONSTERS).sngX) Or _
                        (intNEWX < arrMONSTERS(nCMONSTERS).sngX + arrcMONSTERPICS(arrMONSTERS(nC).intTYPE).width And arrFLAILS(nC).sngX + csprFLAIL.width > arrMONSTERS(nCMONSTERS).sngX) Then
                            If arrFLAILS(nC).didGOTHROUGH(nCMONSTERS) = False Then
                                intDELETEMONSTER = nCMONSTERS
                            End If
                        End If
                    End If
                End If
                nCMONSTERS = nCMONSTERS + 1
            Loop
            
            If intDELETEMONSTER <> -1 Then ' delete a monster
                arrMONSTERS(intDELETEMONSTER).intHEALTH = arrMONSTERS(intDELETEMONSTER).intHEALTH - intFLAILPOWER
                If arrMONSTERS(intDELETEMONSTER).intHEALTH < 1 Then
                    arrMONSTERS(intDELETEMONSTER).bACTIVE = False
                    intMONSTERSKILLED = intMONSTERSKILLED + 1
                    If onlineMODE = False Then lLEVELMONEY = safeADDLONG(lLEVELMONEY, cmontypeMONSTERINFO(arrMONSTERS(intDELETEMONSTER).intTYPE).intMONEYADDEDKILL)
                Else
                    If intFLAILGOTHROUGH > 1 Then arrFLAILS(nC).addGOTHROUGH intDELETEMONSTER
                    If onlineMODE = False Then lLEVELMONEY = safeADDLONG(lLEVELMONEY, cmontypeMONSTERINFO(arrMONSTERS(intDELETEMONSTER).intTYPE).intMONEYADDEDHIT)
                End If
                If arrFLAILS(nC).intGOTHROUGH > 1 Then
                    arrFLAILS(nC).intGOTHROUGH = arrFLAILS(nC).intGOTHROUGH - 1
                Else
                    arrFLAILS(nC).bACTIVE = False
                End If
            End If
            
            arrFLAILS(nC).sngX = intNEWX
            arrFLAILS(nC).sngY = intNEWY
            
            If arrFLAILS(nC).sngX + flailSIZEPX < 0 Or arrFLAILS(nC).sngX > windowX Or arrFLAILS(nC).sngY < -1000 Or arrFLAILS(nC).sngY > windowY - 50 - flailSIZEPX Then
                arrFLAILS(nC).bACTIVE = False
            End If
        End If
        nC = nC + 1
    Loop
    
    If onlineMODE = False And intMONSTERSKILLED + intMONSTERSATTACKEDCASTLE > UBound(arrTOBEMONSTERS) Then
        bEXIT = True
    End If
End Sub

Sub drawEVERYTHING()
Dim nC As Integer
Dim nCMAX As Integer

' draw background
'BitBlt frmATTACK.hDC, 0, 0, frmATTACK.Width, frmATTACK.Height, picBACKGROUND.hDC, intBGLEFT, intBGTOP, vbSrcCopy
BitBlt cbitBUFFER.hdc, 0, 0, cbitBACKGROUND.width, cbitBACKGROUND.height, cbitBACKGROUND.hdc, 0, 0, vbSrcCopy

'draw castle
If lCASTLECURRENTHEALTH > 0 Then
    nC = (csprCASTLE.numberOfFrames - 1) \ (lCASTLEMAXHEALTH / lCASTLECURRENTHEALTH)
Else
    nC = 0
End If

BitBlt cbitBUFFER.hdc, (windowX - csprCASTLE.width) \ 2, castleTOPMARGIN, csprCASTLE.width, csprCASTLE.height, csprCASTLE.frameMaskhDC(nC), 0, 0, vbSrcAnd
BitBlt cbitBUFFER.hdc, (windowX - csprCASTLE.width) \ 2, castleTOPMARGIN, csprCASTLE.width, csprCASTLE.height, csprCASTLE.framehDC(nC), 0, 0, vbSrcPaint

If bEXIT = False Then
    ' draw monsters
    nC = 0
    nCMAX = UBound(arrMONSTERS)
    Do While nC <= nCMAX
        If arrMONSTERS(nC).bACTIVE = True Then
            If arrMONSTERS(nC).sngMOVINGH >= 0 Then
                BitBlt cbitBUFFER.hdc, arrMONSTERS(nC).sngX, arrMONSTERS(nC).sngY, arrcMONSTERPICS(arrMONSTERS(nC).intTYPE).width, arrcMONSTERPICS(arrMONSTERS(nC).intTYPE).height, arrcMONSTERPICS(arrMONSTERS(nC).intTYPE).frameMaskhDC(arrMONSTERS(nC).currentFRAME), 0, 0, vbSrcAnd
                BitBlt cbitBUFFER.hdc, arrMONSTERS(nC).sngX, arrMONSTERS(nC).sngY, arrcMONSTERPICS(arrMONSTERS(nC).intTYPE).width, arrcMONSTERPICS(arrMONSTERS(nC).intTYPE).height, arrcMONSTERPICS(arrMONSTERS(nC).intTYPE).framehDC(arrMONSTERS(nC).currentFRAME), 0, 0, vbSrcPaint
            Else
                BitBlt cbitBUFFER.hdc, arrMONSTERS(nC).sngX, arrMONSTERS(nC).sngY, arrcMONSTERLPICS(arrMONSTERS(nC).intTYPE).width, arrcMONSTERLPICS(arrMONSTERS(nC).intTYPE).height, arrcMONSTERLPICS(arrMONSTERS(nC).intTYPE).frameMaskhDC(arrMONSTERS(nC).currentFRAME), 0, 0, vbSrcAnd
                BitBlt cbitBUFFER.hdc, arrMONSTERS(nC).sngX, arrMONSTERS(nC).sngY, arrcMONSTERLPICS(arrMONSTERS(nC).intTYPE).width, arrcMONSTERLPICS(arrMONSTERS(nC).intTYPE).height, arrcMONSTERLPICS(arrMONSTERS(nC).intTYPE).framehDC(arrMONSTERS(nC).currentFRAME), 0, 0, vbSrcPaint
            End If
            arrMONSTERS(nC).nextFRAME
        End If
        nC = nC + 1
    Loop
    
    ' draw arrows
    nC = 0
    nCMAX = UBound(arrFLAILS)
    Do While nC <= nCMAX
        If arrFLAILS(nC).bACTIVE = True Then
            BitBlt cbitBUFFER.hdc, arrFLAILS(nC).sngX, arrFLAILS(nC).sngY, csprFLAIL.width, csprFLAIL.height, csprFLAIL.frameMaskhDC(arrFLAILS(nC).lCURRENTANIFRAME), 0, 0, vbSrcAnd
            BitBlt cbitBUFFER.hdc, arrFLAILS(nC).sngX, arrFLAILS(nC).sngY, csprFLAIL.width, csprFLAIL.height, csprFLAIL.framehDC(arrFLAILS(nC).lCURRENTANIFRAME), 0, 0, vbSrcPaint
            arrFLAILS(nC).nextFRAME
        End If
        nC = nC + 1
    Loop
End If
    
'draw health
If lCASTLECURRENTHEALTH >= 0 Then
    BitBlt cbitBUFFER.hdc, 10, windowY - cbitHEALTH.height - 20, 30 + ((cbitHEALTH.width - 30) * (lCASTLECURRENTHEALTH / lCASTLEMAXHEALTH)), cbitHEALTH.height, cbitHEALTH.hdc, 0, 0, vbSrcCopy
Else
    BitBlt cbitBUFFER.hdc, 10, windowY - cbitHEALTH.height - 20, 30, cbitHEALTH.height, cbitHEALTH.hdc, 0, 0, vbSrcCopy
End If

drawBUFFER
If lLEVELMONEY <> 0 Then
    lblSCORE.Caption = "Score: " & lLEVELMONEY & "0"
Else
    lblSCORE.Caption = "Score: 0"
End If
End Sub

Sub checkWINLOOSE()
    'check for win/loose
    If bEXIT = True Then
        timerMAIN.Enabled = False
        lineAIM.Visible = False
        If lCASTLECURRENTHEALTH <= 0 Then
            lCASTLECURRENTHEALTH = 0 ' reset health
            If lLEVELMONEY <> 0 Then
                MsgBox "Your castle has fallen! At least you got to keep half of your loot, $" & lLEVELMONEY \ 2 & "0."
            Else
                MsgBox "Your castle has fallen!"
            End If
            lMONEY = safeADDLONG(lMONEY, lLEVELMONEY \ 2)
        Else
            If lLEVELMONEY <> 0 Then
                MsgBox "You beat this level!" & vbCrLf & "You got $" & lLEVELMONEY & "0, plus a level bonus of $" & lCURRENTLEVEL * 2 & "00!"
            Else
                MsgBox "You beat the level!" & vbCrLf & "You got a level bonus of $" & lCURRENTLEVEL * 2 & "00!"
            End If
            lMONEY = safeADDLONG(lMONEY, lLEVELMONEY)
            lMONEY = safeADDLONG(lMONEY, (lCURRENTLEVEL * 20))
            If lLEVEL = lCURRENTLEVEL Then lLEVEL = lLEVEL + 1
        End If
        'frmNEWGAME.Show
        If onlineMODE = False Then
            frmLEVELSELECT.Show ' single player, go to menu
        Else
            frmLOBBY.Show
            'TODO: show store
        End If
        Unload frmATTACK
    End If
End Sub

Private Sub Form_Load()
    Dim bLOADED As Boolean
    bLOADED = True
    
    ' load images
    bLOADED = bLOADED And cbitBACKGROUND.loadFILE(imagePATH & "background.bmp")
    bLOADED = bLOADED And csprCASTLE.loadFRAMES(imagePATH & "castle.bmp", 211, 226, False, True)
    
    bLOADED = bLOADED And cbitHEALTH.loadFILE(imagePATH & "health.bmp")
    
    bLOADED = bLOADED And cbitBUFFER.createNewImage(windowX, windowY)
    
    If bLOADED = False Then
        MsgBox "Error loading images!"
        End
    End If
    
    ' set vars
    Dim nC As Integer
    bEXIT = False
    bFORCEEXIT = False
    lLEVELMONEY = 0
    intMONSTERSKILLED = 0
    intMONSTERSATTACKEDCASTLE = 0
    intCURRENTMONSTER = 0
    lMONSTERSPAWNCOOLDOWN = 0
    ReDim arrTOBEMONSTERS(0 To 0)
    If onlineMODE = False Then ' single player
        sngMOVESPEED = 1 + (lCURRENTLEVEL / 10)
    Else
        sngMOVESPEED = getMOVESPEED
    End If
    
    ReDim arrMONSTERS(0 To 99)
    nC = 0
    Do While nC <= UBound(arrMONSTERS)
        Set arrMONSTERS(nC) = New clsMONSTER
        arrMONSTERS(nC).bACTIVE = False
        nC = nC + 1
    Loop
    
    ReDim arrFLAILS(0 To 99)
    nC = 0
    Do While nC <= UBound(arrFLAILS)
        Set arrFLAILS(nC) = New clsFLAIL
        arrFLAILS(nC).bACTIVE = False
        nC = nC + 1
    Loop
    
    If onlineMODE = False Then
        Dim intTOTALMONSTERS As Integer
        intTOTALMONSTERS = 0
        nC = 0
        Dim nC2 As Integer
        Do While nC < numberOfMonsters
            intTOTALMONSTERS = intTOTALMONSTERS + intMONSTERSONLEVEL(nC)
            nC2 = 0
            If intTOTALMONSTERS <> 0 Then ReDim Preserve arrTOBEMONSTERS(0 To intTOTALMONSTERS - 1)
        
            Do While nC2 < intMONSTERSONLEVEL(nC)
                arrTOBEMONSTERS(intTOTALMONSTERS - 1 - nC2) = nC
                nC2 = nC2 + 1
            Loop
            nC = nC + 1
        Loop
        
        Dim intTEMPSPOT As Integer
        Dim intTEMP As Integer
        nC = 0
        Do While nC < intTOTALMONSTERS - 2 ' -2 is to get to second last monster in array, keeping last monster at last spot
            intTEMPSPOT = Int(Rnd() * (intTOTALMONSTERS - 2)) ' get another random spot
            ' switch the 2 spots
            intTEMP = arrTOBEMONSTERS(nC)
            arrTOBEMONSTERS(nC) = arrTOBEMONSTERS(intTEMPSPOT)
            arrTOBEMONSTERS(intTEMPSPOT) = intTEMP
            nC = nC + 1 ' next spot
        Loop
    End If
    
    frmATTACK.width = (windowX + (frmATTACK.width / Screen.TwipsPerPixelX) - frmATTACK.ScaleWidth) * Screen.TwipsPerPixelX ' width = width + border
    frmATTACK.height = (windowY + (frmATTACK.height / Screen.TwipsPerPixelY) - frmATTACK.ScaleHeight) * Screen.TwipsPerPixelY ' height = height + border
    
    'timerMAIN.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bFORCEEXIT = True ' stop game loop if still running
End Sub
