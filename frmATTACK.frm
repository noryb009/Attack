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
   Begin VB.Timer timerSTART 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   600
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
Dim csprFONT As New clsSPRITE ' font
Dim cbitHEALTH As New clsBITMAP ' health bar
Dim cbitMONHEALTH As New clsBITMAP ' monster health bar

Dim strNEWCHATMSG As String ' new chat message

Dim lSTARTINGHEALTH As Long ' starting health, used for endless mode to restore health after the game

Dim bENDLESSMODE As Boolean ' true if in endless mode

Const keepX = 343 ' X location of flail starting point
Const keepY = 180 ' Y location of flail starting point

Const castleTOPMARGIN = 150 ' space above top of castle image

Sub playGAME()
    Dim currSTARTTIME As Currency ' starting time
    Dim currCURRENTTIME As Currency ' current time
    'Dim currFREQUENCY As Currency ' frame frequency
    Dim dblTIMEBETWEENFRAMES As Double ' time between frames
    
    QueryPerformanceFrequency currCURRENTTIME ' currFREQUENCY ' get the frequency of ticks
    dblTIMEBETWEENFRAMES = currCURRENTTIME / FPS ' currFREQUENCY / FPS ' get time between frames needed to reach FPS
    
    Do While bEXIT = False And bFORCEEXIT = False ' if not exiting yet
        QueryPerformanceCounter currCURRENTTIME ' get current time
        If currCURRENTTIME >= currSTARTTIME + dblTIMEBETWEENFRAMES Then ' if start time + time between frame = current time, then time for the next frame
            QueryPerformanceCounter currSTARTTIME ' store current time as new start time
            moveEVERYTHING ' move everything
            drawEVERYTHING ' draw everything
        Else
            Sleep 1 ' sleep
        End If
        DoEvents ' do any events needed to be done
    Loop
    
    If lCASTLECURRENTHEALTH <= 0 Then ' if you died
        lCASTLECURRENTHEALTH = 0 ' reset health
    End If
    
    If bFORCEEXIT = False And onlineMODE = False Then ' if program not exiting and not online
        drawEVERYTHING ' draw everything a final time
        lineAIM.Visible = False ' hide aim line
        If bENDLESSMODE = True Then ' if in endless mode
            lMONEY = safeADDLONG(lMONEY, lLEVELMONEY) ' add your score to your money
            lCASTLECURRENTHEALTH = lSTARTINGHEALTH ' restore starting health
            
            If lHIGHSCORE < lLEVELMONEY Then ' beat old high score
                MsgBox "You died! You killed " & lMONSTERSKILLED & " monsters, and got a score of " & lLEVELMONEY & "0! You beat your old high score!" ' alert user about their stats
                
                lHIGHSCORE = lLEVELMONEY ' update new high score
                
                Dim dbSAVEFILES As Database ' database link
                Dim recsetSAVES As Recordset ' record set
                
                Set dbSAVEFILES = OpenDatabase(strDATABASEPATH) ' open database
                
                Set recsetSAVES = dbSAVEFILES.OpenRecordset("SELECT * FROM `SaveGames` WHERE `Name`='" & escapeQUOTES(strNAME) & "'") ' get all rows with current username
                
                If recsetSAVES.RecordCount <> 0 Then ' if row exists
                    ' update the highscore in the save row
                    dbSAVEFILES.Execute "UPDATE `SaveGames` SET `Highscore`=" & lHIGHSCORE & " WHERE `Name`='" & escapeQUOTES(strNAME) & "'"
                End If
                
                Set recsetSAVES = Nothing ' close the recordset
                Set dbSAVEFILES = Nothing ' close the database link
            Else
                If lLEVELMONEY <> 0 Then ' if user has money
                    MsgBox "You died! You killed " & lMONSTERSKILLED & " monsters, and got a score of " & lLEVELMONEY & "0! You did not beat your old high score." ' alert user about their stats
                Else
                    MsgBox "You died! You killed " & lMONSTERSKILLED & " monsters, and got a score of 0! You did not beat your old high score." ' alert user about their stats, ($0 not $00)
                End If
            End If
            frmHIGHSCORES.Show ' show high scores form
            frmHIGHSCORES.strWHEREISBACK = "levelSelect" ' go to the level select form when done
        ElseIf lCASTLECURRENTHEALTH <= 0 Then ' if you died
            If lLEVELMONEY > 1 Then ' if you have money (1\2 rounds down to 0)
                MsgBox "Your castle has fallen! You keep half of your money for this level, $" & lLEVELMONEY \ 2 & "0." ' alert user they keep half of their money
                lMONEY = safeADDLONG(lMONEY, lLEVELMONEY \ 2) ' add half your money
            Else ' you don't have any money
                MsgBox "Your castle has fallen!" ' alert user that they lost
            End If
            frmLEVELSELECT.Show ' go to level selection menu
        Else ' you won
            If lLEVELMONEY <> 0 Then ' if you have money
                MsgBox "You beat this level!" & vbCrLf & "You got $" & lLEVELMONEY & "0, plus a level bonus of $" & lCURRENTLEVEL * 2 & "00!" ' alert user they keep their money, plus a level bonus
                lMONEY = safeADDLONG(lMONEY, lLEVELMONEY) ' add your money
            Else ' you don't have any money
                MsgBox "You beat the level!" & vbCrLf & "You got a level bonus of $" & lCURRENTLEVEL * 2 & "00!" ' alert user that they won and give them a level bonus
            End If
            lMONEY = safeADDLONG(lMONEY, (lCURRENTLEVEL * 20)) ' add level bonus
            If lLEVEL = lCURRENTLEVEL Then ' if on latest unlocked level
                lLEVEL = lLEVEL + 1 ' unlock next level
            End If
            frmLEVELSELECT.Show ' go to level selection menu
        End If
    End If
    
    Unload frmATTACK ' hide this form
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If onlineMODE = True Then ' if online
        If KeyAscii >= vbKeySpace And KeyAscii <= 126 Then ' if in visible ascii key set (" " to "~")
            If Len(strNAME) + 2 + Len(strNEWCHATMSG) < maxLENGTHOFMSGINGAME Then ' enough room for extra char (2 is for Len(": "))
                strNEWCHATMSG = strNEWCHATMSG & Chr$(KeyAscii) ' add the char
            Else ' not enough room for extra char
                Beep ' error sound
            End If
        ElseIf KeyAscii = vbKeyBack Then ' backspace
            If Len(strNEWCHATMSG) > 1 Then ' if more then 1 char
                strNEWCHATMSG = Left$(strNEWCHATMSG, Len(strNEWCHATMSG) - 1) ' remove last char
            Else ' one or less chars
                strNEWCHATMSG = "" ' remove everything
            End If
        ElseIf KeyAscii = vbKeyReturn Then ' user pressed enter, send message
            If Trim(strNEWCHATMSG) <> "" Then ' if you have something written
                cSERVER(0).sendString "chat", strNEWCHATMSG ' send message
                strNEWCHATMSG = "" ' clear message
            End If
        End If
    End If
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) ' user moved mouse
    If Button = 1 Or Button = 3 Or Button = 5 Or Button = 7 Then ' if left mouse button down
        lineAIM.X2 = x ' set line second point's X as mouse's current X location
        lineAIM.Y2 = y ' set line second point's Y as mouse's current Y location
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) ' user lifted mouse button
    Const divideSPEED = 10 ' pixels mouse has to move for flail to move 1 pixel each tick
    
    If Button = 1 Then ' mouse released
        lineAIM.Visible = False ' make aim line invisible
        
        If (lineAIM.X1 - lineAIM.X2) \ divideSPEED = 0 And (lineAIM.Y1 - lineAIM.Y2) \ divideSPEED = 0 Then ' if user clicked, didn't move mouse much
            Exit Sub ' exit
        End If
        
        Dim nC As Integer
        Dim nC2 As Integer
        nC = 0
        nC2 = 0
        Do While nC2 < intFLAILAMOUNT ' for each flail amount (upgrade)
            If onlineMODE = False Then ' if not online
                Do While nC <= lFLAILARRAYSIZE ' for each flail spot
                    If arrFLAILS(nC).bACTIVE = False Then ' if flail spot isn't used
                        arrFLAILS(nC).bACTIVE = True ' use flail spot
                        arrFLAILS(nC).sngX = keepX ' starting X location
                        arrFLAILS(nC).sngY = keepY ' starting Y location
                        arrFLAILS(nC).sngMOVINGV = (lineAIM.Y1 - lineAIM.Y2) \ divideSPEED ' vertical speed
                        arrFLAILS(nC).sngMOVINGH = (lineAIM.X1 - lineAIM.X2) \ divideSPEED ' horizontal speed
                        
                        arrFLAILS(nC).sngMOVINGV = arrFLAILS(nC).sngMOVINGV + (((intFLAILAMOUNT / 2) - 0.5 - nC2) * 4) ' spread flails if multiple
                        'arrFLAILS(nC).sngMOVINGH = arrFLAILS(nC).sngMOVINGH + (((intFLAILAMOUNT / 2) - 0.5 - nC2) * 2)
                        
                        arrFLAILS(nC).lOWNER = -1 ' single player flail colour
                        
                        arrFLAILS(nC).intGOTHROUGH = intFLAILGOTHROUGH ' set go through left
                        arrFLAILS(nC).clearWENTTHROUGH ' clear the list of monsters that flail went through
                        Exit Do ' found a flail spot
                    End If
                    nC = nC + 1 ' next flail spot
                Loop
            Else
                cSERVER(0).sendString "newFla", True & "~" & keepX & "~" & keepY & "~" & _
                (lineAIM.Y1 - lineAIM.Y2) \ divideSPEED + (((intFLAILAMOUNT / 2) - 0.5 - nC2) * 4) & "~" & _
                (lineAIM.X1 - lineAIM.X2) \ divideSPEED & "~" & _
                intFLAILGOTHROUGH & "~" & _
                True ' tell server to make new flail
            End If
            nC2 = nC2 + 1 ' next flail in flail amount
        Loop
    End If
End Sub

Sub drawBUFFER() ' draw buffer to screen
    frmATTACK.Cls ' clear screen
    
    'BitBlt frmATTACK.hdc, 0, 0, cbitBUFFER.width, cbitBUFFER.height, cbitBUFFER.hdc, 0, 0, vbSrcCopy
    StretchBlt frmATTACK.hdc, 0, 0, frmATTACK.ScaleWidth, frmATTACK.ScaleHeight, cbitBUFFER.hdc, 0, 0, cbitBUFFER.width, cbitBUFFER.height, vbSrcCopy ' copy buffer to screen
    
    frmATTACK.Refresh ' refresh screen
End Sub

Function widthOFTEXT(ByRef strTEXT) As Long
    widthOFTEXT = Len(strTEXT) * csprFONT.width
End Function

Sub writeONIMAGE(ByVal strTEXT As String, lDESTDC As Long, ByVal x As Long, y As Long, Optional lMAXWIDTH As Long = -1)
    If lMAXWIDTH <> -1 And lMAXWIDTH < Len(strTEXT) Then ' if string can't fit in spot
        If lMAXWIDTH > 3 Then ' if enough room for "..."
            strTEXT = Left$(strTEXT, lMAXWIDTH - 3) ' get the most amount of text that can fit in the spot
            strTEXT = strTEXT + "..." ' add "..." to the end
        Else ' not enough room for "..."
            strTEXT = Left(strTEXT, lMAXWIDTH) ' get the most amount of text that can fit in the spot
        End If
    End If
    
    Dim nC As Integer
    nC = 0
    Dim lTEXTLEN As Long ' length of strTEXT
    lTEXTLEN = Len(strTEXT)
    Dim intCURRENTCHAR As Integer ' current letter in integer form
    Do While nC < lTEXTLEN
        intCURRENTCHAR = Asc(Mid$(strTEXT, nC + 1, 1)) ' get the character number of the current letter
        If intCURRENTCHAR > 127 Then ' out of ascii key set
            intCURRENTCHAR = 0 ' set to 0, nothing
        End If
        BitBlt lDESTDC, x, y, csprFONT.width, csprFONT.height, csprFONT.frameMaskhDC(intCURRENTCHAR), 0, 0, vbSrcAnd
        BitBlt lDESTDC, x, y, csprFONT.width, csprFONT.height, csprFONT.framehDC(intCURRENTCHAR), 0, 0, vbSrcPaint
        x = x + csprFONT.width
        nC = nC + 1
    Loop
End Sub

Sub spawnMONSTER() ' spawn a monster
    Dim intMONSTERTYPE As Integer
    If bENDLESSMODE = True Then ' if in endless mode
        ' gradually add monsters
        If (lCURRENTMONSTER \ 5) > numberOfMonsters Then ' if user has unlocked all monsters
            intMONSTERTYPE = Int(Rnd() * numberOfMonsters) ' pick a random monster
        Else ' user hasn't unlocked all the monsters yet
            intMONSTERTYPE = Int(Rnd() * (lCURRENTMONSTER \ 5)) ' pick a random unlocked monster
        End If
    Else ' not in endless mode
        intMONSTERTYPE = arrTOBEMONSTERS(lCURRENTMONSTER) ' get monster type from monster waiting list
    End If
    nC = 0
    Do While nC < lMONSTERARRAYSIZE ' for each monster spot
        If arrMONSTERS(nC).bACTIVE = False Then ' if monster spot not used
            arrMONSTERS(nC).bACTIVE = True ' use monster spot
            arrMONSTERS(nC).intTYPE = intMONSTERTYPE ' set monster type
            arrMONSTERS(nC).currentFRAME = 0 ' reset frame counter
            
            arrMONSTERS(nC).sngX = Int(Rnd() * 2) ' random starting side
            If arrMONSTERS(nC).sngX = 0 Then ' if on left side
                arrMONSTERS(nC).sngX = 0 - arrcMONSTERPICS(intMONSTERTYPE).width ' start at left side
                arrMONSTERS(nC).sngMOVINGH = 1 * cmontypeMONSTERINFO(intMONSTERTYPE).sngXSPEED ' go right
            Else ' on right side
                arrMONSTERS(nC).sngX = windowX ' start at right side
                arrMONSTERS(nC).sngMOVINGH = -1 * cmontypeMONSTERINFO(intMONSTERTYPE).sngXSPEED ' go left
            End If
            arrMONSTERS(nC).sngY = cmontypeMONSTERINFO(intMONSTERTYPE).intSTARTINGY ' set starting Y location
            arrMONSTERS(nC).sngMOVINGV = cmontypeMONSTERINFO(intMONSTERTYPE).sngYSPEED ' set vertical going down speed
            arrMONSTERS(nC).intHEALTH = cmontypeMONSTERINFO(intMONSTERTYPE).intMAXHEALTH ' set starting health
            
            lCURRENTMONSTER = safeADDLONG(lCURRENTMONSTER, 1) ' one more monster
            Exit Do ' found spot, exit
        End If
        nC = nC + 1 ' next monster spot
    Loop
End Sub

Sub moveEVERYTHING() ' move all the monsters and flails
    Dim nC As Long
    
    If bENDLESSMODE = True Then ' if endless mode
        sngMOVESPEED = (lCURRENTMONSTER / 40) + 1 ' update moveing speed multiplier
    End If
    
    ' spawn monsters
    If onlineMODE = False Then
        If lMONSTERSPAWNCOOLDOWN = 0 Then ' if it has been a while since last spawn
            Dim bSPAWN As Boolean
            bSPAWN = False ' default: don't spawn
            If bENDLESSMODE = True Then
                If lCURRENTMONSTER <= lMONSTERSKILLED + lMONSTERSATTACKEDCASTLE Then ' force if nobody on screen
                    bSPAWN = True ' spawn
                ElseIf Int(Rnd() * 200) < (lCURRENTMONSTER / 2) Then ' randomly
                    bSPAWN = True ' spawn
                End If
            ElseIf lCURRENTMONSTER <= UBound(arrTOBEMONSTERS) Then ' normal mode, if there are monsters waiting to be spawned
                If lCURRENTMONSTER <= lMONSTERSKILLED + lMONSTERSATTACKEDCASTLE + (lCURRENTLEVEL \ 3) Then ' force if not many on screen
                    bSPAWN = True ' spawn
                ElseIf Int(Rnd() * 200) < lCURRENTLEVEL And lCURRENTMONSTER <= UBound(arrTOBEMONSTERS) Then ' randomly if some monsters are waiting
                    bSPAWN = True ' spawn
                End If
            End If
            If bSPAWN = True Then ' if going to spawn
                spawnMONSTER ' spawn the monster
                lMONSTERSPAWNCOOLDOWN = 20 ' wait a bit for the next monster
            End If
        Else
            lMONSTERSPAWNCOOLDOWN = lMONSTERSPAWNCOOLDOWN - 1 ' count down cooldown time
        End If
    End If
    
    ' move monsters
    nC = 0
    Do While nC < lMONSTERARRAYSIZE ' for each monster
        If arrMONSTERS(nC).bACTIVE = True Then ' if monster is active
            arrMONSTERS(nC).moveMONSTER ' move monster
        End If
        nC = nC + 1 ' next monster
    Loop
    
    If lCASTLECURRENTHEALTH <= 0 And onlineMODE = False Then ' if dead and not online
        bEXIT = True ' exit
    End If
    
    ' move flails
    nC = 0
    Do While nC < lFLAILARRAYSIZE ' for each flail
        If arrFLAILS(nC).bACTIVE = True Then ' if flail is active
            If onlineMODE = True Then ' if online
                arrFLAILS(nC).moveFLAIL ' move the flail
            Else ' offline
                lLEVELMONEY = safeADDLONG(lLEVELMONEY, arrFLAILS(nC).moveFLAIL) ' move the flail and add to your score
            End If
        End If
        nC = nC + 1 ' next flail
    Loop
    
    If onlineMODE = False And bENDLESSMODE = False And lMONSTERSKILLED + lMONSTERSATTACKEDCASTLE > UBound(arrTOBEMONSTERS) Then ' if offline, not endless mode, and you have defeated all the monsters
        bEXIT = True ' exit
    End If
End Sub

Sub drawEVERYTHING() ' draw everything to the screen
    Dim nC As Integer
    
    ' draw background
    BitBlt cbitBUFFER.hdc, 0, 0, cbitBACKGROUND.width, cbitBACKGROUND.height, cbitBACKGROUND.hdc, 0, 0, vbSrcCopy
    
    ' draw castle
    If lCASTLECURRENTHEALTH > 0 Then ' if you still have health
        nC = (csprCASTLE.numberOfFrames - 1) \ (lCASTLEMAXHEALTH / lCASTLECURRENTHEALTH) ' use castle image that is closest to your health level
    Else ' if you are dead
        nC = 0 ' use dead castle image
    End If
    
    ' draw castle
    BitBlt cbitBUFFER.hdc, (windowX - csprCASTLE.width) \ 2, castleTOPMARGIN, csprCASTLE.width, csprCASTLE.height, csprCASTLE.frameMaskhDC(nC), 0, 0, vbSrcAnd ' draw castle mask
    BitBlt cbitBUFFER.hdc, (windowX - csprCASTLE.width) \ 2, castleTOPMARGIN, csprCASTLE.width, csprCASTLE.height, csprCASTLE.framehDC(nC), 0, 0, vbSrcPaint ' draw castle
    
    If bEXIT = False Then ' if not exiting
        ' draw monsters
        nC = 0
        Do While nC < lMONSTERARRAYSIZE ' for each monster
            If arrMONSTERS(nC).bACTIVE = True Then ' if monster is active
                If arrMONSTERS(nC).sngMOVINGH >= 0 Then ' if moving right
                    BitBlt cbitBUFFER.hdc, arrMONSTERS(nC).sngX, arrMONSTERS(nC).sngY, arrcMONSTERPICS(arrMONSTERS(nC).intTYPE).width, arrcMONSTERPICS(arrMONSTERS(nC).intTYPE).height, arrcMONSTERPICS(arrMONSTERS(nC).intTYPE).frameMaskhDC(arrMONSTERS(nC).currentFRAME), 0, 0, vbSrcAnd ' draw right monster mask
                    BitBlt cbitBUFFER.hdc, arrMONSTERS(nC).sngX, arrMONSTERS(nC).sngY, arrcMONSTERPICS(arrMONSTERS(nC).intTYPE).width, arrcMONSTERPICS(arrMONSTERS(nC).intTYPE).height, arrcMONSTERPICS(arrMONSTERS(nC).intTYPE).framehDC(arrMONSTERS(nC).currentFRAME), 0, 0, vbSrcPaint ' draw right monster
                Else
                    BitBlt cbitBUFFER.hdc, arrMONSTERS(nC).sngX, arrMONSTERS(nC).sngY, arrcMONSTERLPICS(arrMONSTERS(nC).intTYPE).width, arrcMONSTERLPICS(arrMONSTERS(nC).intTYPE).height, arrcMONSTERLPICS(arrMONSTERS(nC).intTYPE).frameMaskhDC(arrMONSTERS(nC).currentFRAME), 0, 0, vbSrcAnd ' draw left monster mask
                    BitBlt cbitBUFFER.hdc, arrMONSTERS(nC).sngX, arrMONSTERS(nC).sngY, arrcMONSTERLPICS(arrMONSTERS(nC).intTYPE).width, arrcMONSTERLPICS(arrMONSTERS(nC).intTYPE).height, arrcMONSTERLPICS(arrMONSTERS(nC).intTYPE).framehDC(arrMONSTERS(nC).currentFRAME), 0, 0, vbSrcPaint ' draw left monster
                End If
                arrMONSTERS(nC).nextFRAME ' go to the next frame
                
                ' monster health
                If arrMONSTERS(nC).intHEALTH <> cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).intMAXHEALTH And arrMONSTERS(nC).intHEALTH > 0 Then ' if monster has been attacked
                    BitBlt cbitBUFFER.hdc, arrMONSTERS(nC).sngX + ((arrcMONSTERLPICS(arrMONSTERS(nC).intTYPE).width - cbitMONHEALTH.width) \ 2), arrMONSTERS(nC).sngY - 10, cbitMONHEALTH.width / (cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).intMAXHEALTH / arrMONSTERS(nC).intHEALTH), cbitMONHEALTH.height, cbitMONHEALTH.hdc, 0, 0, vbSrcCopy ' draw health bar
                End If
            End If
            nC = nC + 1 ' next monster
        Loop
        
        ' draw arrows
        nC = 0
        Do While nC < lFLAILARRAYSIZE ' for each flail
            If arrFLAILS(nC).bACTIVE = True Then ' if flail is active
                BitBlt cbitBUFFER.hdc, Int(arrFLAILS(nC).sngX), Int(arrFLAILS(nC).sngY), csprFLAIL.width, csprFLAIL.height, csprFLAIL.frameMaskhDC(arrFLAILS(nC).lOWNER + 1), 0, 0, vbSrcAnd ' draw flail mask
                BitBlt cbitBUFFER.hdc, Int(arrFLAILS(nC).sngX), Int(arrFLAILS(nC).sngY), csprFLAIL.width, csprFLAIL.height, csprFLAIL.framehDC(arrFLAILS(nC).lOWNER + 1), 0, 0, vbSrcPaint ' draw flail
            End If
            nC = nC + 1 ' next flail
        Loop
    End If
    
    ' draw your username
    writeONIMAGE strNAME, cbitBUFFER.hdc, 10, 455
    
    ' draw health
    writeONIMAGE "Health", cbitBUFFER.hdc, 10, 470
    If lCASTLECURRENTHEALTH > 0 Then ' if you still have health
        BitBlt cbitBUFFER.hdc, 60, 470, cbitHEALTH.width * (lCASTLECURRENTHEALTH / lCASTLEMAXHEALTH), cbitHEALTH.height, cbitHEALTH.hdc, 0, 0, vbSrcCopy ' display health
        writeONIMAGE lCASTLECURRENTHEALTH & "0/" & lCASTLEMAXHEALTH & "0", cbitBUFFER.hdc, (windowX - widthOFTEXT(lCASTLECURRENTHEALTH & "0/" & lCASTLEMAXHEALTH & "0")) \ 2, 485
    Else
        writeONIMAGE "0/" & lCASTLEMAXHEALTH & "0", cbitBUFFER.hdc, windowX - 10 - widthOFTEXT("0/" & lCASTLEMAXHEALTH & "0"), 485
    End If
    
    ' draw monsters left
    If bENDLESSMODE = True Then ' if in endless mode
        writeONIMAGE "Monsters defeated: " & CStr(lMONSTERSKILLED), cbitBUFFER.hdc, (windowX - widthOFTEXT("Monsters defeated: " & CStr(lMONSTERSKILLED))) \ 2, 455 ' draw monsters killed
    ElseIf onlineMODE = True Then ' if online
        writeONIMAGE "Monsters left: " & CStr(lMONSTERSLEFT), cbitBUFFER.hdc, (windowX - widthOFTEXT("Monsters left: " & CStr(lMONSTERSLEFT))) \ 2, 455 ' draw monsters left
    Else ' offline, but not in endless mode
        writeONIMAGE "Monsters left: " & getMONSTERSLEFT, cbitBUFFER.hdc, (windowX - widthOFTEXT("Monsters left: " & getMONSTERSLEFT)) \ 2, 455 ' draw monsters left
    End If
    
    ' draw score
    If onlineMODE = False Then ' if offline
        If lLEVELMONEY = 0 Then ' if no score
            writeONIMAGE "Score: 0", cbitBUFFER.hdc, windowX - widthOFTEXT("Score: 0") - 10, 455 ' display your score (0, not 00)
        Else
            writeONIMAGE "Score: " & lLEVELMONEY & "0", cbitBUFFER.hdc, windowX - widthOFTEXT("Score: " & lLEVELMONEY & "0") - 10, 455 ' display score
        End If
    Else ' online
        Dim intCURRENTPLAYER As Integer
        intCURRENTPLAYER = 0
        nC = 0
        Do While nC < MAXCLIENTS ' for each clients
            If ccinfoPLAYERINFO(nC).strNAME <> "" Then ' if player spot is being used
                writeONIMAGE ccinfoPLAYERINFO(nC).strNAME, cbitBUFFER.hdc, ((windowX \ (intPLAYERS + 1)) * (intCURRENTPLAYER + 1)) - (widthOFTEXT(ccinfoPLAYERINFO(nC).strNAME) \ 2), 5, (windowX \ intPLAYERS + 1 \ csprFONT.width) ' draw player name
                If ccinfoPLAYERINFO(nC).lLEVELSCORE = 0 Then ' if no score
                    writeONIMAGE "0", cbitBUFFER.hdc, ((windowX \ (intPLAYERS + 1)) * (intCURRENTPLAYER + 1)) - (widthOFTEXT("0") \ 2), 19, -1 ' draw player score
                Else
                    writeONIMAGE CStr(ccinfoPLAYERINFO(nC).lLEVELSCORE) & "0", cbitBUFFER.hdc, ((windowX \ (intPLAYERS + 1)) * (intCURRENTPLAYER + 1)) - (widthOFTEXT(CStr(ccinfoPLAYERINFO(nC).lLEVELSCORE) & "0") \ 2), 19, (windowX \ intPLAYERS + 1 \ csprFONT.width) ' draw player score
                End If
                intCURRENTPLAYER = intCURRENTPLAYER + 1 ' found one more player
            End If
            nC = nC + 1 ' next client
        Loop
    End If
    
    ' draw chat log
    If onlineMODE = True Then ' if online
        If strNEWCHATMSG <> "" Then ' if typing
            nC = 0 ' show all messages in the chat log
        Else ' not typing
            nC = UBound(strCHATLOG) - 2 ' show last 3 messages in chat log
        End If
        Do While nC <= UBound(strCHATLOG)
            If strCHATLOG(nC) <> "" Then ' if not empty
                writeONIMAGE strCHATLOG(nC), cbitBUFFER.hdc, 5, 200 + (nC * csprFONT.height), maxLENGTHOFMSGINGAME ' write message to screen
            End If
            nC = nC + 1
        Loop
        ' draw message you are currently editing
        If strNEWCHATMSG <> "" Then
            writeONIMAGE strNAME & ": " & strNEWCHATMSG, cbitBUFFER.hdc, 5, 200 + (UBound(strCHATLOG) * csprFONT.height) + csprFONT.height, -1 ' write new message to screen
        End If
    End If
    
    drawBUFFER ' draw buffer to the screen
End Sub

Private Sub Form_Load()
    Dim bLOADED As Boolean
    bLOADED = True
    
    ' load images
    bLOADED = bLOADED And cbitBACKGROUND.loadFILE(strIMAGEPATH & "background.bmp")
    bLOADED = bLOADED And csprCASTLE.loadFRAMES(strIMAGEPATH & "castle.bmp", 211, 226, False, True)
    
    bLOADED = bLOADED And cbitHEALTH.loadFILE(strIMAGEPATH & "health.bmp")
    bLOADED = bLOADED And cbitMONHEALTH.loadFILE(strIMAGEPATH & "monHealth.bmp")
    
    bLOADED = bLOADED And csprFONT.loadFRAMES(strIMAGEPATH & "font.bmp", 7, 14, False, True)
    If csprFONT.numberOfFrames <> 128 Then ' if wrong number of frames
        bLOADED = False ' error
    End If
    
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
    lMONSTERSKILLED = 0
    lMONSTERSATTACKEDCASTLE = 0
    lCURRENTMONSTER = 0
    lMONSTERSPAWNCOOLDOWN = 0
    strNEWCHATMSG = ""
    lSTARTINGHEALTH = lCASTLECURRENTHEALTH
    ReDim arrTOBEMONSTERS(0 To 0)
    If onlineMODE = False Then ' single player
        sngMOVESPEED = 1 + (lCURRENTLEVEL / 10)
    Else
        sngMOVESPEED = getMOVESPEED
    End If
    
    nC = 0
    Do While nC < lMONSTERARRAYSIZE
        arrMONSTERS(nC).bACTIVE = False
        nC = nC + 1
    Loop
    
    nC = 0
    Do While nC < lFLAILARRAYSIZE
        arrFLAILS(nC).bACTIVE = False
        nC = nC + 1
    Loop
    
    If onlineMODE = False Then
        ' check if in endless mode
        If lCURRENTLEVEL = 11 Then ' if on level 11 (endless level)
            bENDLESSMODE = True ' you are in endless mode
        Else
            bENDLESSMODE = False ' you are not in endless mode
            
            ' count monsters
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
            
            ' randomize monster order (but keep last monster in place)
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
    Else
        bENDLESSMODE = False ' you are not in endless mode
    End If
    
    frmATTACK.width = (windowX + (frmATTACK.width / Screen.TwipsPerPixelX) - frmATTACK.ScaleWidth) * Screen.TwipsPerPixelX ' width = width + border
    frmATTACK.height = (windowY + (frmATTACK.height / Screen.TwipsPerPixelY) - frmATTACK.ScaleHeight) * Screen.TwipsPerPixelY ' height = height + border
    
    timerSTART.Enabled = True ' start game after timer ticks
End Sub

Private Sub Form_Unload(Cancel As Integer) ' on form close
    bFORCEEXIT = True ' stop game loop if still running
End Sub

Private Sub timerSTART_Timer()
    timerSTART.Enabled = False ' disable timer
    playGAME ' play the game
End Sub
