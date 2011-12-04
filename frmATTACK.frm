VERSION 5.00
Begin VB.Form frmATTACK 
   AutoRedraw      =   -1  'True
   Caption         =   "Attack"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   Picture         =   "frmATTACK.frx":0000
   ScaleHeight     =   306
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   421
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMONSTERL 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   165
      Index           =   6
      Left            =   4680
      Picture         =   "frmATTACK.frx":B272
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   29
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picMONSTERBACKL 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   165
      Index           =   6
      Left            =   5160
      Picture         =   "frmATTACK.frx":B2BC
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   28
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picMONSTERBACK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   165
      Index           =   6
      Left            =   3720
      Picture         =   "frmATTACK.frx":B306
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   27
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picMONSTER 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   165
      Index           =   6
      Left            =   3240
      Picture         =   "frmATTACK.frx":B351
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   26
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picMONSTER 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   5
      Left            =   3240
      Picture         =   "frmATTACK.frx":B39C
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   25
      Top             =   3120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox picMONSTERBACK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   5
      Left            =   3720
      Picture         =   "frmATTACK.frx":B44D
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   24
      Top             =   3120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox picMONSTERBACKL 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   5
      Left            =   5160
      Picture         =   "frmATTACK.frx":B4FE
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   23
      Top             =   3120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox picMONSTERL 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   5
      Left            =   4680
      Picture         =   "frmATTACK.frx":B5AE
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   22
      Top             =   3120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox picMONSTERL 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   4680
      Picture         =   "frmATTACK.frx":B65E
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   21
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picMONSTERBACKL 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   5280
      Picture         =   "frmATTACK.frx":B703
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   20
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picMONSTERBACKL 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   1
      Left            =   5280
      Picture         =   "frmATTACK.frx":B7A8
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   19
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picMONSTERL 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   1
      Left            =   4680
      Picture         =   "frmATTACK.frx":B84D
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picMONSTERL 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   225
      Index           =   2
      Left            =   4680
      Picture         =   "frmATTACK.frx":BBB8
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   17
      Top             =   1440
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picMONSTERBACKL 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   225
      Index           =   2
      Left            =   5280
      Picture         =   "frmATTACK.frx":BC05
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picMONSTERBACKL 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   3
      Left            =   5280
      Picture         =   "frmATTACK.frx":BC52
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picMONSTERL 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   3
      Left            =   4680
      Picture         =   "frmATTACK.frx":C02F
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picMONSTERL 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   480
      Index           =   4
      Left            =   4680
      Picture         =   "frmATTACK.frx":C490
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   13
      Top             =   2520
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox picMONSTERBACKL 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   480
      Index           =   4
      Left            =   5280
      Picture         =   "frmATTACK.frx":C82F
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox picMONSTERBACK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   480
      Index           =   4
      Left            =   3840
      Picture         =   "frmATTACK.frx":CBDC
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   11
      Top             =   2520
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox picMONSTER 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   480
      Index           =   4
      Left            =   3240
      Picture         =   "frmATTACK.frx":CF89
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox picMONSTER 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   3
      Left            =   3240
      Picture         =   "frmATTACK.frx":D328
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picMONSTERBACK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   3
      Left            =   3840
      Picture         =   "frmATTACK.frx":D789
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picMONSTERBACK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   225
      Index           =   2
      Left            =   3840
      Picture         =   "frmATTACK.frx":DB66
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picMONSTER 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   225
      Index           =   2
      Left            =   3240
      Picture         =   "frmATTACK.frx":DBB0
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picMONSTER 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   1
      Left            =   3240
      Picture         =   "frmATTACK.frx":DBFA
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picMONSTERBACK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   1
      Left            =   3840
      Picture         =   "frmATTACK.frx":DF65
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picFLAILBACK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   270
      Left            =   720
      Picture         =   "frmATTACK.frx":E00A
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   3
      Top             =   3000
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox picFLAIL 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   270
      Left            =   360
      Picture         =   "frmATTACK.frx":E063
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   2
      Top             =   3000
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox picMONSTERBACK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   3840
      Picture         =   "frmATTACK.frx":E0BC
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picMONSTER 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   3240
      Picture         =   "frmATTACK.frx":E161
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer timerMAIN 
      Interval        =   10
      Left            =   1920
      Top             =   2040
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
      TabIndex        =   30
      Top             =   120
      Width           =   1575
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
'Dim bISPRESSED_UP As Boolean
'Dim bISPRESSED_DOWN As Boolean
'Dim bISPRESSED_LEFT As Boolean
'Dim bISPRESSED_RIGHT As Boolean

'Dim intBGLEFT As Integer
'Dim intBGTOP As Integer

Dim intPLAYERS As Integer

Dim arrMONSTERS(0 To 99) As New clsMONSTER
Dim arrFLAILS(0 To 99) As New clsFLAIL

Dim arrTOBEMONSTERS() As Integer
Dim intCURRENTMONSTER As Integer
Dim intMONSTERSKILLED As Integer

Const landHEIGHT = 336
Const keepX = 338
Const keepY = 190

Const vWindowX = 700
Const vWindowY = 500

Private Sub Form_Load()
Randomize
intPLAYERS = 1
intCURRENTMONSTER = 0
intMONSTERSKILLED = 0

Dim intTOTALMONSTERS As Integer
intTOTALMONSTERS = 0
Dim nC As Integer
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
Do While nC < intTOTALMONSTERS - 1 ' -1 is to keep last monster at last spot
    intTEMPSPOT = Int(Rnd() * intTOTALMONSTERS - 1) + 1
    intTEMP = arrTOBEMONSTERS(nC)
    arrTOBEMONSTERS(nC) = arrTOBEMONSTERS(intTEMPSPOT)
    arrTOBEMONSTERS(intTEMPSPOT) = intTEMP
    nC = nC + 1
Loop
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
    Do While nC < UBound(arrFLAILS)
        If arrFLAILS(nC).bACTIVE = False Then
            arrFLAILS(nC).bACTIVE = True
            arrFLAILS(nC).sngX = keepX
            arrFLAILS(nC).sngY = keepY
            arrFLAILS(nC).sngMOVINGV = (lineAIM.Y1 - lineAIM.Y2) \ divideSPEED
            arrFLAILS(nC).sngMOVINGH = (lineAIM.X1 - lineAIM.X2) \ divideSPEED
            Exit Do
        End If
        nC = nC + 1
    Loop
End If
End Sub

Private Sub Form_Resize()

    frmATTACK.Width = 700 + (frmATTACK.Width - frmATTACK.ScaleWidth)
    frmATTACK.Height = 500 + (frmATTACK.Height - frmATTACK.ScaleHeight)
    Exit Sub
ResizeErr:
    If Err.Number = 384 Then
        Exit Sub ' no error on minimize
    End If
End Sub

Private Sub timerMAIN_Timer()
Const moveSPEED = 1
Const distMOVEBACK = 50

frmATTACK.Cls
' draw background
'BitBlt frmATTACK.hDC, 0, 0, frmATTACK.Width, frmATTACK.Height, picBACKGROUND.hDC, intBGLEFT, intBGTOP, vbSrcCopy

If intMONSTERSKILLED = UBound(arrTOBEMONSTERS) + 1 Then
    MsgBox "You beat this level!"
    frmLEVELSELECT.Show
    Unload frmATTACK
    Exit Sub
End If

Dim nC As Integer

' spawn monsters
If Int(Rnd() * (intPLAYERS * 5 + 100)) < intPLAYERS Or intMONSTERSKILLED = intCURRENTMONSTER Then ' randomly, but force if no monsters currently on screen
    If intCURRENTMONSTER <= UBound(arrTOBEMONSTERS) Then
    
        nC = 0
        Do While nC <= UBound(arrMONSTERS)
            If arrMONSTERS(nC).bACTIVE = False Then
                arrMONSTERS(nC).bACTIVE = True
                arrMONSTERS(nC).intTYPE = arrTOBEMONSTERS(intCURRENTMONSTER) 'Int(Rnd() * numberOfMonsters)
                
                ' set default variables
                arrMONSTERS(nC).sngX = Int(Rnd() * 2)
                If arrMONSTERS(nC).sngX = 0 Then
                    arrMONSTERS(nC).sngX = 0 - picMONSTER(arrMONSTERS(nC).intTYPE).ScaleWidth
                    arrMONSTERS(nC).sngMOVINGH = 1 ' go left
                Else
                    arrMONSTERS(nC).sngX = vWindowX + picMONSTER(arrMONSTERS(nC).intTYPE).ScaleWidth
                    arrMONSTERS(nC).sngMOVINGH = -1 ' go right
                End If
                
                arrMONSTERS(nC).intHEALTH = 1
                arrMONSTERS(nC).sngY = landHEIGHT - picMONSTER(arrMONSTERS(nC).intTYPE).ScaleHeight
                
                ' set height and health for a specific monster
                Select Case arrMONSTERS(nC).intTYPE
                    Case 0
                    Case 1
                        arrMONSTERS(nC).intHEALTH = 2
                    Case 2
                        arrMONSTERS(nC).sngY = landHEIGHT - 200 - picMONSTER(arrMONSTERS(nC).intTYPE).ScaleHeight
                        arrMONSTERS(nC).sngMOVINGH = arrMONSTERS(nC).sngMOVINGH * 2
                    Case 3
                        arrMONSTERS(nC).intHEALTH = 4
                        arrMONSTERS(nC).sngMOVINGH = arrMONSTERS(nC).sngMOVINGH / 5
                    Case 4
                        arrMONSTERS(nC).intHEALTH = 2
                        arrMONSTERS(nC).sngY = landHEIGHT - 300 - picMONSTER(arrMONSTERS(nC).intTYPE).ScaleHeight
                End Select
                Exit Do
            End If
            nC = nC + 1
        Loop
        intCURRENTMONSTER = intCURRENTMONSTER + 1
    End If
End If

' draw monsters

'Dim spMONSTER As New StdPicture
nC = 0
Do While nC <= UBound(arrMONSTERS)
    If arrMONSTERS(nC).bACTIVE = True Then
        arrMONSTERS(nC).sngX = arrMONSTERS(nC).sngX + moveSPEED * arrMONSTERS(nC).sngMOVINGH
        If (arrMONSTERS(nC).sngMOVINGH < 0 And arrMONSTERS(nC).sngX + picMONSTER(arrMONSTERS(nC).intTYPE).ScaleWidth < 0) Or (arrMONSTERS(nC).sngMOVINGH > 0 And arrMONSTERS(nC).sngX > vWindowX) Then
            arrMONSTERS(nC).bACTIVE = False
        Else
            If arrMONSTERS(nC).sngMOVINGH >= 0 Then
                'Set spMONSTER = picMONSTER(arrMONSTERS(nC).intTYPE).Picture
                BitBlt frmATTACK.hDC, arrMONSTERS(nC).sngX, arrMONSTERS(nC).sngY, picMONSTERBACK(arrMONSTERS(nC).intTYPE).ScaleWidth, picMONSTERBACK(arrMONSTERS(nC).intTYPE).ScaleHeight, picMONSTERBACK(arrMONSTERS(nC).intTYPE).hDC, 0, 0, vbSrcAnd
                BitBlt frmATTACK.hDC, arrMONSTERS(nC).sngX, arrMONSTERS(nC).sngY, picMONSTER(arrMONSTERS(nC).intTYPE).ScaleWidth, picMONSTER(arrMONSTERS(nC).intTYPE).ScaleHeight, picMONSTER(arrMONSTERS(nC).intTYPE).hDC, 0, 0, vbSrcPaint
            Else
                'Set spMONSTER = picMONSTERL(arrMONSTERS(nC).intTYPE).Picture
                BitBlt frmATTACK.hDC, arrMONSTERS(nC).sngX, arrMONSTERS(nC).sngY, picMONSTERBACKL(arrMONSTERS(nC).intTYPE).ScaleWidth, picMONSTERBACKL(arrMONSTERS(nC).intTYPE).ScaleHeight, picMONSTERBACKL(arrMONSTERS(nC).intTYPE).hDC, 0, 0, vbSrcAnd
                BitBlt frmATTACK.hDC, arrMONSTERS(nC).sngX, arrMONSTERS(nC).sngY, picMONSTERL(arrMONSTERS(nC).intTYPE).ScaleWidth, picMONSTERL(arrMONSTERS(nC).intTYPE).ScaleHeight, picMONSTERL(arrMONSTERS(nC).intTYPE).hDC, 0, 0, vbSrcPaint
            End If
            'PaintPicture spMONSTER, arrMONSTERS(nC).sngX * (frmATTACK.ScaleWidth / vWindowX), arrMONSTERS(nC).sngY * (frmATTACK.ScaleHeight / vWindowY), picMONSTER(arrMONSTERS(nC).intTYPE).ScaleWidth * (frmATTACK.ScaleWidth / vWindowX), picMONSTER(arrMONSTERS(nC).intTYPE).ScaleHeight * (frmATTACK.ScaleHeight / vWindowY)
        End If
    End If
    nC = nC + 1
Loop

' draw arrows
nC = 0
'Dim spFLAIL As New StdPicture
'Set spFLAIL = picFLAIL.Picture
Do While nC <= UBound(arrFLAILS)
    If arrFLAILS(nC).bACTIVE = True Then
        Dim intNEWX As Integer
        Dim intNEWY As Integer
        Dim intDELETEMONSTER As Integer
        intDELETEMONSTER = -1 ' don't delete a monster
        intNEWX = Int(arrFLAILS(nC).sngX) + arrFLAILS(nC).sngMOVINGH
        intNEWY = Int(arrFLAILS(nC).sngY) + arrFLAILS(nC).sngMOVINGV
        
        arrFLAILS(nC).sngMOVINGV = arrFLAILS(nC).sngMOVINGV + 0.5
        If arrFLAILS(nC).sngMOVINGH < 0 Then
            arrFLAILS(nC).sngMOVINGH = arrFLAILS(nC).sngMOVINGH + 0.1
        Else
            arrFLAILS(nC).sngMOVINGH = arrFLAILS(nC).sngMOVINGH - 0.1
        End If
        
        'check if hitting monster
        Dim nCMONSTERS As Integer
        nCMONSTERS = 0
        Do While nCMONSTERS <= UBound(arrMONSTERS)
            If arrMONSTERS(nCMONSTERS).bACTIVE = True Then
                If (arrFLAILS(nC).sngY < arrMONSTERS(nCMONSTERS).sngY + picMONSTER(arrMONSTERS(nCMONSTERS).intTYPE).Height And intNEWY + picFLAIL.Height > arrMONSTERS(nCMONSTERS).sngY) Or _
                (intNEWY < arrMONSTERS(nCMONSTERS).sngY + picMONSTER(arrMONSTERS(nCMONSTERS).intTYPE).Height And arrFLAILS(nC).sngY + picFLAIL.Height > arrMONSTERS(nCMONSTERS).sngY) Then
                    'If arrFLAILS(nC).sngx + picFLAIL.Width < arrMONSTERS(nCMONSTERS).sngx And intNEWX + picFLAIL.Width > arrMONSTERS(nCMONSTERS).sngx Then
                    If (arrFLAILS(nC).sngX < arrMONSTERS(nCMONSTERS).sngX + picMONSTER(arrMONSTERS(nCMONSTERS).intTYPE).Width And intNEWX + picFLAIL.Width > arrMONSTERS(nCMONSTERS).sngX) Or _
                    (intNEWX < arrMONSTERS(nCMONSTERS).sngX + picMONSTER(arrMONSTERS(nCMONSTERS).intTYPE).Width And arrFLAILS(nC).sngX + picFLAIL.Width > arrMONSTERS(nCMONSTERS).sngX) Then
                        intDELETEMONSTER = nCMONSTERS
                    End If
                End If
            End If
            nCMONSTERS = nCMONSTERS + 1
        Loop
        
        If intDELETEMONSTER <> -1 Then ' delete a monster
            arrMONSTERS(intDELETEMONSTER).intHEALTH = arrMONSTERS(intDELETEMONSTER).intHEALTH - intFLAILPOWER - 1
            If arrMONSTERS(intDELETEMONSTER).intHEALTH < 1 Then
                arrMONSTERS(intDELETEMONSTER).bACTIVE = False
                intMONSTERSKILLED = intMONSTERSKILLED + 1
                intSCORE = intSCORE + (arrMONSTERS(intDELETEMONSTER).intTYPE * 100) + 100
            Else
                intSCORE = intSCORE + (arrMONSTERS(intDELETEMONSTER).intTYPE * 10) + 10
            End If
            arrFLAILS(nC).bACTIVE = False
        End If
        
        arrFLAILS(nC).sngX = intNEWX
        arrFLAILS(nC).sngY = intNEWY
        
        
        If arrFLAILS(nC).sngX + picFLAIL.ScaleWidth < 0 Or arrFLAILS(nC).sngX > vWindowX Or arrFLAILS(nC).sngY < -1000 Or arrFLAILS(nC).sngY > vWindowY Then
            arrFLAILS(nC).bACTIVE = False
        Else
            BitBlt frmATTACK.hDC, arrFLAILS(nC).sngX, arrFLAILS(nC).sngY, picFLAILBACK.ScaleWidth, picFLAILBACK.ScaleHeight, picFLAILBACK.hDC, 0, 0, vbSrcAnd
            BitBlt frmATTACK.hDC, arrFLAILS(nC).sngX, arrFLAILS(nC).sngY, picFLAIL.ScaleWidth, picFLAIL.ScaleHeight, picFLAIL.hDC, 0, 0, vbSrcPaint
            'PaintPicture spFLAIL, arrFLAILS(nC).sngX * (frmATTACK.ScaleWidth / vWindowX), arrFLAILS(nC).sngY * (frmATTACK.ScaleHeight / vWindowY), picFLAIL.ScaleWidth * (frmATTACK.ScaleWidth / vWindowX), picFLAIL.ScaleHeight * (frmATTACK.ScaleHeight / vWindowY)
        End If
    End If
    nC = nC + 1
Loop

frmATTACK.Refresh

lblSCORE.Caption = "Score: " & intSCORE
End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Select Case KeyCode
'    Case vbKeyUp
'        bISPRESSED_UP = True
'    Case vbKeyDown
'        bISPRESSED_DOWN = True
'    Case vbKeyLeft
'        bISPRESSED_LEFT = True
''    Case vbKeyRight
'        bISPRESSED_RIGHT = True
'End Select
'End Sub

'Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'Select Case KeyCode
'    Case vbKeyUp
'        bISPRESSED_UP = False
'    Case vbKeyDown
'        bISPRESSED_DOWN = False
'    Case vbKeyLeft
'        bISPRESSED_LEFT = False
'    Case vbKeyRight
'        bISPRESSED_RIGHT = False
'End Select
'End Sub
