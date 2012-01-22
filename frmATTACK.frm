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
   WindowState     =   2  'Maximized
   Begin VB.Timer timerSTART 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   600
   End
   Begin VB.Line lineAIM 
      Visible         =   0   'False
      X1              =   8
      X2              =   8
      Y1              =   32
      Y2              =   64
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

Dim strNEWCHATMSG As String ' new chat message
Dim lSTARTINGHEALTH As Long ' starting health, used for endless mode to restore health after the game
Dim bENDLESSMODE As Boolean ' true if in endless mode
Dim intCANNONDIRECTION As Integer ' direction of cannon

Dim bPAUSED As Boolean ' if true, game is paused
Dim bAIMING As Boolean ' if true, user is aiming

Const castleTOPMARGIN = 150 ' space above top of castle image

Const keepX = 343 ' X location of flail starting point
Const keepY = 180 ' Y location of flail starting point

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
            If bPAUSED = False Then ' if not paused
                moveEVERYTHING ' move everything
                drawEVERYTHING ' draw everything
            End If
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
                MsgBox "You died! You killed " & lMONSTERSKILLED & " monsters, and got a score of " & lLEVELMONEY & "0! You beat your old high score!", vbOKOnly, programNAME ' alert user about their stats
                
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
                MsgBox "You died! You killed " & lMONSTERSKILLED & " monsters, and got a score of " & addZEROIFNOTZERO(lLEVELMONEY) & "! You did not beat your old high score.", vbOKOnly, programNAME ' alert user about their stats
            End If
            frmHIGHSCORES.Show ' show high scores form
            frmHIGHSCORES.strWHEREISBACK = "levelSelect" ' go to the level select form when done
        ElseIf lCASTLECURRENTHEALTH <= 0 Then ' if you died
            If lLEVELMONEY > 1 Then ' if you have money (1\2 rounds down to 0)
                MsgBox "Your castle has fallen! You keep half of your money for this level, $" & lLEVELMONEY \ 2 & "0.", vbOKOnly, programNAME ' alert user they keep half of their money
                lMONEY = safeADDLONG(lMONEY, lLEVELMONEY \ 2) ' add half your money
            Else ' you don't have any money
                MsgBox "Your castle has fallen!", vbOKOnly, programNAME ' alert user that they lost
            End If
            frmLEVELSELECT.Show ' go to level selection menu
        Else ' you won
            If lLEVELMONEY <> 0 Then ' if you have money
                MsgBox "You beat this level!" & vbCrLf & "You got $" & lLEVELMONEY & "0, plus a level bonus of $" & lCURRENTLEVEL * 2 & "00!", vbOKOnly, programNAME ' alert user they keep their money, plus a level bonus
                lMONEY = safeADDLONG(lMONEY, lLEVELMONEY) ' add your money
            Else ' you don't have any money
                MsgBox "You beat the level!" & vbCrLf & "You got a level bonus of $" & lCURRENTLEVEL * 2 & "00!", vbOKOnly, programNAME ' alert user that they won and give them a level bonus
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
    Else ' offline
        If KeyAscii = vbKeyP Or KeyAscii = vbKeySpace Or KeyAscii = 112 Then ' pause/unpause the game (p, P, or space)
            If bPAUSED = True Then ' if game is currently paused
                bPAUSED = False ' unpause the game
            Else
                bPAUSED = True ' pause the game
                bAIMING = False
                lineAIM.Visible = False
                drawEVERYTHING ' draw everything to show the pause game message
            End If
        ElseIf KeyAscii = vbKeyEscape And bPAUSED = True Then ' if exiting
            bFORCEEXIT = True ' exit loop
            frmLEVELSELECT.Show ' show level select
        End If
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 And bPAUSED = False Then ' if left click and not paused
        bAIMING = True ' currently aiming
        lineAIM.Visible = True ' show the aim line
        lineAIM.X1 = x ' set the starting x
        lineAIM.Y1 = y ' set the starting y
        lineAIM.X2 = x ' set the current x
        lineAIM.Y2 = y ' set the current y
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) ' user moved mouse
    If Button = 1 Or Button = 3 Or Button = 5 Or Button = 7 Then ' if left mouse button down
        If bAIMING = True Then ' if currently aiming
            lineAIM.X2 = x ' set line second point's X as mouse's current X location
            lineAIM.Y2 = y ' set line second point's Y as mouse's current Y location
        End If
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) ' user lifted mouse button
    If bAIMING = False Then ' if not aiming
        Exit Sub ' exit
    End If
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
                DoEvents
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

Function widthOFTEXT(ByRef strTEXT As String) As Long
    widthOFTEXT = Len(strTEXT) * csprFONT.width
End Function

Sub writeONIMAGE(ByVal strTEXT As String, lDESTDC As Long, ByVal x As Long, y As Long, Optional intALIGNMENT As Integer = -1, Optional lMAXWIDTH As Long = -1)
    If lMAXWIDTH <> -1 And lMAXWIDTH < Len(strTEXT) Then ' if string can't fit in spot
        If lMAXWIDTH > 3 Then ' if enough room for "..."
            strTEXT = Left$(strTEXT, lMAXWIDTH - 3) ' get the most amount of text that can fit in the spot
            strTEXT = strTEXT + "..." ' add "..." to the end
        Else ' not enough room for "..."
            strTEXT = Left(strTEXT, lMAXWIDTH) ' get the most amount of text that can fit in the spot
        End If
    End If
    
    ' intALIGNMENT: -1=left, 0=center, 1=right
    If intALIGNMENT = 1 Then ' text at right
        x = x - (widthOFTEXT(strTEXT)) ' remove width of text, text ends at x location
    ElseIf intALIGNMENT = 0 Then ' text in center
        x = x - (widthOFTEXT(strTEXT) \ 2) ' remove half of width of text, text middle is at x location
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
            If bENDLESSMODE = True Then ' TODO: + part of monsters
                If lCURRENTMONSTER <= lMONSTERSKILLED + lMONSTERSATTACKEDCASTLE + (lCURRENTMONSTER \ 5) Then ' force if nobody on screen
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
                If bENDLESSMODE = True Then
                    If (lCURRENTMONSTER \ 20) >= 20 Then
                        lMONSTERSPAWNCOOLDOWN = 1 ' wait a bit for the next monster
                    Else
                        lMONSTERSPAWNCOOLDOWN = 20 - (lCURRENTMONSTER \ 20) ' wait a bit for the next monster
                    End If
                Else
                    lMONSTERSPAWNCOOLDOWN = 20
                End If
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
        'nC = csprCASTLE.numberOfFrames - ((csprCASTLE.numberOfFrames - 1) \ (lCASTLEMAXHEALTH / lCASTLECURRENTHEALTH)) ' use castle image that is closest to your health level
        nC = csprCASTLE.numberOfFrames - (lCASTLECURRENTHEALTH / (lCASTLEMAXHEALTH / csprCASTLE.numberOfFrames))
    Else ' if you are dead
        nC = csprCASTLE.numberOfFrames - 1 ' use dead castle image
    End If
    
    ' above code isn't perfect
    If nC < 0 Then ' if less then 0
        nC = 0 ' show castle at lowest health
    ElseIf nC >= csprCASTLE.numberOfFrames Then ' if above last frame
        nC = csprCASTLE.numberOfFrames - 1 ' show castle at full health
    End If
    
    ' draw castle
    BitBlt cbitBUFFER.hdc, (windowX - csprCASTLE.width) \ 2, castleTOPMARGIN, csprCASTLE.width, csprCASTLE.height, csprCASTLE.frameMaskhDC(nC), 0, 0, vbSrcAnd ' draw castle mask
    BitBlt cbitBUFFER.hdc, (windowX - csprCASTLE.width) \ 2, castleTOPMARGIN, csprCASTLE.width, csprCASTLE.height, csprCASTLE.framehDC(nC), 0, 0, vbSrcPaint ' draw castle
    
    ' draw cannon
    Dim lXDIFF As Long
    lXDIFF = Abs(lineAIM.X2 - lineAIM.X1)
    Dim lYDIFF As Long
    lYDIFF = Abs(lineAIM.Y2 - lineAIM.Y1)
    
    If bAIMING = True And lXDIFF <> 0 And lYDIFF <> 0 Then ' if aiming and you have moved the mouse
        If lineAIM.X1 >= lineAIM.X2 And lineAIM.Y1 < lineAIM.Y2 Then ' top right
            If lXDIFF * 2 < lYDIFF Then
                intCANNONDIRECTION = 0 ' North
            ElseIf lXDIFF < lYDIFF * 2 Then
                intCANNONDIRECTION = 1 ' North-East
            Else
                intCANNONDIRECTION = 2 ' East
            End If
        ElseIf lineAIM.X1 > lineAIM.X2 And lineAIM.Y1 >= lineAIM.Y2 Then ' bottom right
            If lXDIFF > lYDIFF * 2 Then
                intCANNONDIRECTION = 2 ' East
            ElseIf lXDIFF * 2 > lYDIFF Then
                intCANNONDIRECTION = 3 ' South-East
            Else
                intCANNONDIRECTION = 4 ' South
            End If
        ElseIf lineAIM.X1 <= lineAIM.X2 And lineAIM.Y1 > lineAIM.Y2 Then ' bottom left
            If lXDIFF * 2 < lYDIFF Then
                intCANNONDIRECTION = 4 ' South
            ElseIf lXDIFF < lYDIFF * 2 Then
                intCANNONDIRECTION = 5 ' South-West
            Else
                intCANNONDIRECTION = 6 ' West
            End If
        Else ' top left
           If lXDIFF > lYDIFF * 2 Then
                intCANNONDIRECTION = 6 ' West
            ElseIf lXDIFF * 2 > lYDIFF Then
                intCANNONDIRECTION = 7 ' North-West
            Else
                intCANNONDIRECTION = 0 ' North
            End If
        End If
    End If
    
    BitBlt cbitBUFFER.hdc, (windowX - csprCANNON.width) \ 2, castleTOPMARGIN + 25, csprCANNON.width, csprCANNON.height, csprCANNON.frameMaskhDC(intCANNONDIRECTION), 0, 0, vbSrcAnd ' draw cannon mask
    BitBlt cbitBUFFER.hdc, (windowX - csprCANNON.width) \ 2, castleTOPMARGIN + 25, csprCANNON.width, csprCANNON.height, csprCANNON.framehDC(intCANNONDIRECTION), 0, 0, vbSrcPaint ' draw cannon
    
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
        
        ' draw flails
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
    End If
    writeONIMAGE addZEROIFNOTZERO(lCASTLECURRENTHEALTH) & "/" & lCASTLEMAXHEALTH & "0", cbitBUFFER.hdc, windowX \ 2, 485, 0 ' write numbers in center of screen
    
    ' draw monsters left
    If bENDLESSMODE = True Then ' if in endless mode
        writeONIMAGE "Monsters defeated: " & CStr(lMONSTERSKILLED), cbitBUFFER.hdc, windowX \ 2, 455, 0 ' draw monsters killed in center of screen
    ElseIf onlineMODE = True Then ' if online
        writeONIMAGE "Monsters left: " & CStr(lMONSTERSLEFT), cbitBUFFER.hdc, windowX \ 2, 455, 0 ' draw monsters left in center of screen
    Else ' offline, but not in endless mode
        writeONIMAGE "Monsters left: " & getMONSTERSLEFT, cbitBUFFER.hdc, windowX \ 2, 455, 0 ' draw monsters left in center of screen
    End If
    
    ' draw score
    If onlineMODE = False Then ' if offline
        writeONIMAGE "Score: " & addZEROIFNOTZERO(lLEVELMONEY), cbitBUFFER.hdc, windowX - 10, 455, 1 ' display score on right side
    Else ' online
        Dim intCURRENTPLAYER As Integer
        intCURRENTPLAYER = 0
        nC = 0
        Do While nC < MAXCLIENTS ' for each clients
            If ccinfoPLAYERINFO(nC).strNAME <> "" Then ' if player spot is being used
                writeONIMAGE ccinfoPLAYERINFO(nC).strNAME, cbitBUFFER.hdc, (windowX \ (intPLAYERS + 1)) * (intCURRENTPLAYER + 1), 5, 0, (windowX \ intPLAYERS + 1 \ csprFONT.width) ' draw player name
                writeONIMAGE addZEROIFNOTZERO(ccinfoPLAYERINFO(nC).lLEVELSCORE), cbitBUFFER.hdc, (windowX \ (intPLAYERS + 1)) * (intCURRENTPLAYER + 1), 19, 0, (windowX \ intPLAYERS + 1 \ csprFONT.width) ' draw player score
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
    
    If onlineMODE = False And bPAUSED = True Then ' if paused
        writeONIMAGE "Game paused.", cbitBUFFER.hdc, windowX \ 2, 0, 0, -1
        writeONIMAGE "Press space to resume, or escape to exit.", cbitBUFFER.hdc, windowX \ 2, csprFONT.height, 0, -1
    End If
    
    drawBUFFER ' draw buffer to the screen
End Sub

Private Sub Form_Load()
    ' set vars
    Dim nC As Integer
    bEXIT = False
    bFORCEEXIT = False
    bPAUSED = False
    bAIMING = False
    lLEVELMONEY = 0
    lMONSTERSKILLED = 0
    lMONSTERSATTACKEDCASTLE = 0
    lCURRENTMONSTER = 0
    lMONSTERSPAWNCOOLDOWN = 0
    strNEWCHATMSG = ""
    lSTARTINGHEALTH = lCASTLECURRENTHEALTH
    intCANNONDIRECTION = 0 ' default cannon direction: North
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
