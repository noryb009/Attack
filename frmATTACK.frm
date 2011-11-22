VERSION 5.00
Begin VB.Form frmATTACK 
   AutoRedraw      =   -1  'True
   Caption         =   "Attack"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   270
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   336
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMONSTER 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   390
      Left            =   3360
      Picture         =   "frmATTACK.frx":0000
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Timer timerMAIN 
      Interval        =   10
      Left            =   1920
      Top             =   2040
   End
   Begin VB.PictureBox picBACKGROUND 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   18060
      Left            =   3240
      Picture         =   "frmATTACK.frx":0672
      ScaleHeight     =   1200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1200
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   18060
   End
   Begin VB.Shape shapeYOU 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   2160
      Top             =   1800
      Width           =   375
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

Dim arrMONSTERS(0 To 99) As New clsMONSTER

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

Private Sub timerMAIN_Timer()
Const moveSPEED = 10
Const distMOVEBACK = 50

frmATTACK.Cls
' draw background
BitBlt frmATTACK.hDC, 0, 0, frmATTACK.Width, frmATTACK.Height, picBACKGROUND.hDC, intBGLEFT, intBGTOP, vbSrcCopy

If bISPRESSED_UP = True And bISPRESSED_DOWN = False Then
    If intBGTOP - moveSPEED < 0 Then
        If intBGTOP <> 0 Then intBGTOP = 0
    Else
        intBGTOP = intBGTOP - moveSPEED
    End If
ElseIf bISPRESSED_DOWN = True And bISPRESSED_UP = False Then
    If frmATTACK.ScaleHeight + intBGTOP + moveSPEED > picBACKGROUND.ScaleHeight Then
        If intBGTOP <> picBACKGROUND.ScaleHeight - frmATTACK.ScaleHeight Then intBGTOP = picBACKGROUND.ScaleHeight - frmATTACK.ScaleHeight
    Else
        intBGTOP = intBGTOP + moveSPEED
    End If
End If
If bISPRESSED_LEFT = True And bISPRESSED_RIGHT = False Then
    If intBGLEFT - moveSPEED < 0 Then
        If intBGLEFT <> 0 Then intBGLEFT = 0
    Else
        intBGLEFT = intBGLEFT - moveSPEED
    End If
ElseIf bISPRESSED_RIGHT = True And bISPRESSED_LEFT = False Then
    If frmATTACK.ScaleWidth + intBGLEFT + moveSPEED > picBACKGROUND.ScaleWidth Then
        If intBGLEFT <> picBACKGROUND.ScaleWidth - frmATTACK.ScaleWidth Then intBGLEFT = picBACKGROUND.ScaleWidth - frmATTACK.ScaleWidth
    Else
        intBGLEFT = intBGLEFT + moveSPEED
    End If
End If

Dim nC As Integer

' spawn monsters
If Int(Rnd() * 50) = 0 Then
       nC = 0
    Do While nC <= UBound(arrMONSTERS)
        If arrMONSTERS(nC).bACTIVE = False Then
            arrMONSTERS(nC).bACTIVE = True
            arrMONSTERS(nC).intX = Int(Rnd() * 1000)
            arrMONSTERS(nC).intY = Int(Rnd() * 1000)
            Exit Do
        End If
        nC = nC + 1
    Loop
End If

nC = 0
Do While nC <= UBound(arrMONSTERS)
    If arrMONSTERS(nC).bACTIVE = True Then
        BitBlt frmATTACK.hDC, arrMONSTERS(nC).intX, arrMONSTERS(nC).intY, picMONSTER.ScaleWidth, picMONSTER.ScaleHeight, picBACKGROUND.hDC, 0, 0, vbSrcCopy
    End If
    nC = nC + 1
Loop

frmATTACK.Refresh
End Sub
