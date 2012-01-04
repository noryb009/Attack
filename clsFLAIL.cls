VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFLAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

Public bACTIVE As Boolean ' if flail is on screen
Public lOWNER As Long ' only used in server: who shot the flail
Public sngX As Single ' X location
Public sngY As Single ' Y location
Public sngMOVINGH As Single ' X moving speed
Public sngMOVINGV As Single ' Y moving speed
Public intGOTHROUGH As Integer ' number of go through left
Dim longWENTTHROUGH(0 To 10) As Long ' array of monsters flail has gone through
Dim lSIZEOFlongWENTTHROUGH As Long ' size of longWENTTHROUGH to not have to call VB functions

Public Sub clearWENTTHROUGH() ' clear the go through
    lSIZEOFlongWENTTHROUGH = 0 ' reset the size tracker
End Sub

Public Function didGOTHROUGH(intCHECK As Integer) As Boolean ' check if flail went through monster
    Dim nC As Integer
    nC = 1 ' skip first array spot, not used
    didGOTHROUGH = False ' default: didn't go through
    Do While nC <= lSIZEOFlongWENTTHROUGH ' while less then size of longWENTTHROUGH
        If longWENTTHROUGH(nC) = intCHECK Then ' if flail went through this monster
            didGOTHROUGH = True ' did go through monster
            Exit Do ' found monster, don't need to keep looking
        End If
        nC = nC + 1 ' next go through
    Loop
End Function

Public Sub addGOTHROUGH(intNEW As Integer) ' add go through to
    lSIZEOFlongWENTTHROUGH = lSIZEOFlongWENTTHROUGH + 1 ' add to size tracker
    If UBound(longWENTTHROUGH) < lSIZEOFlongWENTTHROUGH Then ' if not big enough to hold extra flail
        Exit Sub ' exit
    End If
    longWENTTHROUGH(lSIZEOFlongWENTTHROUGH) = intNEW ' remember new monster
End Sub

Public Function moveFLAIL() As Long ' move the flail
    moveFLAIL = 0 ' no score added yet
    
    Dim lMOVESPLIT As Long
    lMOVESPLIT = Abs(sngMOVINGH \ 5) + Abs(sngMOVINGV \ 5) + 1 ' calculate number of little moves the big move should take
    
    Dim sngNEWX As Integer ' new X location
    Dim sngNEWY As Integer ' new Y location
    Dim lDELETEMONSTER As Long ' monster the flail is hitting
    
    Dim lMOVECOUNT As Long
    lMOVECOUNT = 0
    
    Do While lMOVECOUNT < lMOVESPLIT And bACTIVE = True ' for each small move (and flail is still active)
        sngNEWX = sngX + sngMOVINGH / lMOVESPLIT ' add X distance
        sngNEWY = sngY + (sngMOVINGV / lMOVESPLIT) ' add Y distance
        
        lDELETEMONSTER = -1 ' don't delete a monster
        
        'check if hitting monster
        Dim nC As Integer
        nC = 0
        Do While nC < lMONSTERARRAYSIZE ' for each monster
            If arrMONSTERS(nC).bACTIVE = True Then ' if monster is active
                ' if at same elevation
                If (sngY < arrMONSTERS(nC).sngY + cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).lHEIGHT And sngNEWY + cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).lHEIGHT > arrMONSTERS(nC).sngY) Or _
                (sngNEWY < arrMONSTERS(nC).sngY + cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).lHEIGHT And sngY + flailSIZEPX > arrMONSTERS(nC).sngY) Then
                    ' if same vertical spot
                    If (sngX < arrMONSTERS(nC).sngX + cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).lWIDTH And sngNEWX + flailSIZEPX > arrMONSTERS(nC).sngX) Or _
                    (sngNEWX < arrMONSTERS(nC).sngX + cmontypeMONSTERINFO(arrMONSTERS(nC).intTYPE).lWIDTH And sngX + flailSIZEPX > arrMONSTERS(nC).sngX) Then
                        If didGOTHROUGH(nC) = False Then ' if you haven't gone through monster yet
                            lDELETEMONSTER = nC ' attack this monster
                            Exit Do ' you found a monster, continue
                        End If
                    End If
                End If
            End If
            nC = nC + 1 ' next monster
        Loop
        
        If lDELETEMONSTER <> -1 Then ' if you found a monster to attack, then attack monster
            arrMONSTERS(lDELETEMONSTER).intHEALTH = arrMONSTERS(lDELETEMONSTER).intHEALTH - intFLAILPOWER ' take away health from monster
            If arrMONSTERS(lDELETEMONSTER).intHEALTH < 1 Then ' if monster has no health left
                arrMONSTERS(lDELETEMONSTER).bACTIVE = False ' disable monster
                lMONSTERSKILLED = safeADDLONG(lMONSTERSKILLED, 1) ' add to number of monsters killed
                moveFLAIL = safeADDLONG(moveFLAIL, cmontypeMONSTERINFO(arrMONSTERS(lDELETEMONSTER).intTYPE).intMONEYADDEDKILL) ' add score for killing the monster
                
                If SERVER = True Then ' if server, update clients
                    broadcastMONSTER lDELETEMONSTER
                End If
            Else
                If intFLAILGOTHROUGH > 1 Then ' if flail can go through more monsters after this one
                    addGOTHROUGH CInt(lDELETEMONSTER) ' don't attack this monster again
                End If
                moveFLAIL = safeADDLONG(moveFLAIL, cmontypeMONSTERINFO(arrMONSTERS(lDELETEMONSTER).intTYPE).intMONEYADDEDHIT) ' add score for hitting, but not killing the monster
            End If
            If intGOTHROUGH > 1 Then ' if go through left
                intGOTHROUGH = intGOTHROUGH - 1 ' one less monster to go through
            Else ' no more go through left
                bACTIVE = False ' flail has gone through all the monsters
            End If
        End If
        
        sngX = sngNEWX ' set new X
        sngY = sngNEWY ' set new Y
        
        If sngX + flailSIZEPX < 0 Or sngX > windowX Or sngY < -1000 Or sngY > windowY - 50 - flailSIZEPX Then ' if out of bounds (-1000 lets flail go off top of screen)
            bACTIVE = False ' disable flail
        End If
        
        lMOVECOUNT = lMOVECOUNT + 1 ' next small move
    Loop
    
    'gravity
    sngMOVINGV = sngMOVINGV + 0.5 ' go down more next time
    If sngMOVINGH < 0 Then ' if going left
        sngMOVINGH = sngMOVINGH + 0.1 ' go a little less left next time
    ElseIf sngMOVINGH > 0 Then ' if going right
        sngMOVINGH = sngMOVINGH - 0.1 ' go a little less right next time
    End If
End Function

Private Sub Class_Initialize() ' object init
    bACTIVE = False ' flail not on screen
End Sub
