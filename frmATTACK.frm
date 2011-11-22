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
      Picture         =   "frmATTACK.frx":0000
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
      Left            =   1320
      Top             =   1200
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
    If shapeYOU.Top < distMOVEBACK Then
        If intBGTOP + moveSPEED > 0 Then
            If intBGTOP <> 0 Then intBGTOP = 0
        Else
            intBGTOP = intBGTOP - moveSPEED
        End If
    Else
        shapeYOU.Top = shapeYOU.Top - moveSPEED
    End If
ElseIf bISPRESSED_DOWN = True And bISPRESSED_UP = False Then
    If shapeYOU.Top + shapeYOU.Height > frmATTACK.ScaleHeight - distMOVEBACK Then
        If intBGTOP + picBACKGROUND.Height - moveSPEED < frmATTACK.ScaleHeight Then
            If intBGTOP + picBACKGROUND.Height <> frmATTACK.ScaleHeight Then intBGTOP = frmATTACK.ScaleHeight - picBACKGROUND.Height
        Else
            intBGTOP = intBGTOP + moveSPEED
        End If
    Else
        shapeYOU.Top = shapeYOU.Top + moveSPEED
    End If
End If
If bISPRESSED_LEFT = True And bISPRESSED_RIGHT = False Then
    If shapeYOU.Left < distMOVEBACK Then
        If intBGLEFT + moveSPEED > 0 Then
            If intBGLEFT <> 0 Then intBGLEFT = 0
        Else
            intBGLEFT = intBGLEFT - moveSPEED
        End If
    Else
        shapeYOU.Left = shapeYOU.Left - moveSPEED
    End If
ElseIf bISPRESSED_RIGHT = True And bISPRESSED_LEFT = False Then
    If shapeYOU.Left + shapeYOU.Width > frmATTACK.ScaleWidth - distMOVEBACK Then
        If intBGLEFT + picBACKGROUND.Width - moveSPEED < frmATTACK.ScaleWidth Then
            If intBGLEFT + picBACKGROUND.Width <> frmATTACK.ScaleWidth Then intBGLEFT = frmATTACK.ScaleWidth - picBACKGROUND.Width
        Else
            intBGLEFT = intBGLEFT + moveSPEED
        End If
    Else
        shapeYOU.Left = shapeYOU.Left + moveSPEED
    End If
End If

frmATTACK.Refresh
End Sub
