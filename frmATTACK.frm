VERSION 5.00
Begin VB.Form frmATTACK 
   AutoRedraw      =   -1  'True
   Caption         =   "Attack"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   Picture         =   "frmATTACK.frx":0000
   ScaleHeight     =   270
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   336
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picARROWBACK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   3960
      Picture         =   "frmATTACK.frx":B272
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picARROW 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   3480
      Picture         =   "frmATTACK.frx":B317
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picMONSTERBACK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   3840
      Picture         =   "frmATTACK.frx":B3BC
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picMONSTER 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   3360
      Picture         =   "frmATTACK.frx":B461
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer timerMAIN 
      Interval        =   10
      Left            =   1920
      Top             =   2040
   End
   Begin VB.Line lineAIM 
      Visible         =   0   'False
      X1              =   48
      X2              =   184
      Y1              =   32
      Y2              =   32
   End
End
Attribute VB_Name = "frmATTACK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bISPRESSED_UP As Boolean
Dim bISPRESSED_DOWN As Boolean
Dim bISPRESSED_LEFT As Boolean
Dim bISPRESSED_RIGHT As Boolean

Dim intBGLEFT As Integer
Dim intBGTOP As Integer

Dim intPLAYERS As Integer

Dim arrMONSTERS(0 To 99) As New clsMONSTER
Dim arrARROWS(0 To 99) As New clsARROW

Const landHEIGHT = 336
Const keepX = 370
Const keepY = 200

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyUp
        bISPRESSED_UP = True
    Case vbKeyDown
        bISPRESSED_DOWN = True
    Case vbKeyLeft
        bISPRESSED_LEFT = True
    Case vbKeyRight
        bISPRESSED_RIGHT = True
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyUp
        bISPRESSED_UP = False
    Case vbKeyDown
        bISPRESSED_DOWN = False
    Case vbKeyLeft
        bISPRESSED_LEFT = False
    Case vbKeyRight
        bISPRESSED_RIGHT = False
End Select
End Sub

Private Sub Form_Load()
Randomize
intPLAYERS = 1
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    lineAIM.Visible = True
    lineAIM.X1 = X
    lineAIM.Y1 = Y
    lineAIM.X2 = X
    lineAIM.Y2 = Y
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Or Button = 3 Or Button = 5 Or Button = 7 Then
    lineAIM.X2 = X
    lineAIM.Y2 = Y
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const divideSPEED = 10

If Button = 1 Then
    lineAIM.Visible = False
    If (lineAIM.X1 - lineAIM.X2) \ divideSPEED = 0 And (lineAIM.Y1 - lineAIM.Y2) \ divideSPEED = 0 Then
        Exit Sub
    End If
    
    Dim nC As Integer
    nC = 0
    Do While nC < UBound(arrARROWS)
        If arrARROWS(nC).bACTIVE = False Then
            arrARROWS(nC).bACTIVE = True
            arrARROWS(nC).intX = keepX
            arrARROWS(nC).intY = keepY
            arrARROWS(nC).sMOVINGV = (lineAIM.Y1 - lineAIM.Y2) \ divideSPEED
            arrARROWS(nC).sMOVINGH = (lineAIM.X1 - lineAIM.X2) \ divideSPEED
            Exit Do
        End If
        nC = nC + 1
    Loop
End If
End Sub

Private Sub Form_Resize()
frmATTACK.Width = 700 + (frmATTACK.Width - frmATTACK.ScaleWidth)
frmATTACK.Height = 500 + (frmATTACK.Height - frmATTACK.ScaleHeight)
End Sub

Private Sub timerMAIN_Timer()
Const moveSPEED = 1
Const distMOVEBACK = 50

frmATTACK.Cls
' draw background
'BitBlt frmATTACK.hDC, 0, 0, frmATTACK.Width, frmATTACK.Height, picBACKGROUND.hDC, intBGLEFT, intBGTOP, vbSrcCopy

Dim nC As Integer

' spawn monsters
If Int(Rnd() * (intPLAYERS * 5 + 100)) < (intPLAYERS * 2) Then
    nC = 0
    Do While nC <= UBound(arrMONSTERS)
        If arrMONSTERS(nC).bACTIVE = False Then
            arrMONSTERS(nC).bACTIVE = True
            arrMONSTERS(nC).intY = landHEIGHT - picMONSTER.ScaleHeight
            
            arrMONSTERS(nC).intX = Int(Rnd() * 2)
            If arrMONSTERS(nC).intX = 0 Then
                arrMONSTERS(nC).intX = 0 - picMONSTER.ScaleWidth
                arrMONSTERS(nC).sMOVINGH = 1 ' go left
            Else
                arrMONSTERS(nC).intX = frmATTACK.ScaleWidth + picMONSTER.ScaleWidth
                arrMONSTERS(nC).sMOVINGH = -1 ' go right
            End If
            Exit Do
        End If
        nC = nC + 1
    Loop
End If

' draw monsters
nC = 0
Do While nC <= UBound(arrMONSTERS)
    If arrMONSTERS(nC).bACTIVE = True Then
        arrMONSTERS(nC).intX = arrMONSTERS(nC).intX + moveSPEED * arrMONSTERS(nC).sMOVINGH
        If (arrMONSTERS(nC).sMOVINGH < 0 And arrMONSTERS(nC).intX + picMONSTER.ScaleWidth < 0) Or (arrMONSTERS(nC).sMOVINGH > 0 And arrMONSTERS(nC).intX > frmATTACK.ScaleWidth) Then
            arrMONSTERS(nC).bACTIVE = False
        Else
            BitBlt frmATTACK.hDC, arrMONSTERS(nC).intX, arrMONSTERS(nC).intY, picMONSTERBACK.ScaleWidth, picMONSTERBACK.ScaleHeight, picMONSTERBACK.hDC, 0, 0, vbSrcAnd
            BitBlt frmATTACK.hDC, arrMONSTERS(nC).intX, arrMONSTERS(nC).intY, picMONSTER.ScaleWidth, picMONSTER.ScaleHeight, picMONSTER.hDC, 0, 0, vbSrcPaint
        End If
    End If
    nC = nC + 1
Loop

' draw arrows
nC = 0
Do While nC <= UBound(arrARROWS)
    If arrARROWS(nC).bACTIVE = True Then
        Dim intNEWX As Integer
        Dim intNEWY As Integer
        Dim intDELETEMONSTER As Integer
        intDELETEMONSTER = -1 ' don't delete a monster
        intNEWX = arrARROWS(nC).intX + arrARROWS(nC).sMOVINGH
        intNEWY = arrARROWS(nC).intY + arrARROWS(nC).sMOVINGV
        
        arrARROWS(nC).sMOVINGV = arrARROWS(nC).sMOVINGV + 0.5
        If arrARROWS(nC).sMOVINGH < 0 Then
            arrARROWS(nC).sMOVINGH = arrARROWS(nC).sMOVINGH + 0.1
        Else
            arrARROWS(nC).sMOVINGH = arrARROWS(nC).sMOVINGH - 0.1
        End If
        
        'check if hitting monster
        Dim nCMONSTERS As Integer
        nCMONSTERS = 0
        Do While nCMONSTERS <= UBound(arrMONSTERS)
            If arrMONSTERS(nCMONSTERS).bACTIVE = True Then
                If (arrARROWS(nC).intY < arrMONSTERS(nCMONSTERS).intY + picMONSTER.Height And intNEWY + picARROW.Height > arrMONSTERS(nCMONSTERS).intY) Or _
                (intNEWY < arrMONSTERS(nCMONSTERS).intY + picMONSTER.Height And arrARROWS(nC).intY + picARROW.Height > arrMONSTERS(nCMONSTERS).intY) Then
                    'If arrARROWS(nC).intX + picARROW.Width < arrMONSTERS(nCMONSTERS).intX And intNEWX + picARROW.Width > arrMONSTERS(nCMONSTERS).intX Then
                    If (arrARROWS(nC).intX < arrMONSTERS(nCMONSTERS).intX + picMONSTER.Width And intNEWX + picARROW.Width > arrMONSTERS(nCMONSTERS).intX) Or _
                    (intNEWX < arrMONSTERS(nCMONSTERS).intX + picMONSTER.Width And arrARROWS(nC).intX + picARROW.Width > arrMONSTERS(nCMONSTERS).intX) Then
                        intDELETEMONSTER = nCMONSTERS
                    End If
                End If
            End If
            nCMONSTERS = nCMONSTERS + 1
        Loop
        
        If intDELETEMONSTER <> -1 Then ' delete a monster
            arrMONSTERS(intDELETEMONSTER).bACTIVE = False
            arrARROWS(nC).bACTIVE = False
        End If
        
        arrARROWS(nC).intX = intNEWX
        arrARROWS(nC).intY = intNEWY
        
        
        If arrARROWS(nC).intX + picARROW.ScaleWidth < 0 Or arrARROWS(nC).intX > frmATTACK.ScaleWidth Or arrARROWS(nC).intY < -1000 Or arrARROWS(nC).intY > frmATTACK.ScaleHeight Then
            arrARROWS(nC).bACTIVE = False
        Else
            BitBlt frmATTACK.hDC, arrARROWS(nC).intX, arrARROWS(nC).intY, picARROWBACK.ScaleWidth, picARROWBACK.ScaleHeight, picARROWBACK.hDC, 0, 0, vbSrcAnd
            BitBlt frmATTACK.hDC, arrARROWS(nC).intX, arrARROWS(nC).intY, picARROW.ScaleWidth, picARROW.ScaleHeight, picARROW.hDC, 0, 0, vbSrcPaint
        End If
    End If
    nC = nC + 1
Loop

frmATTACK.Refresh
End Sub
