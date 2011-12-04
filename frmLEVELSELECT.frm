VERSION 5.00
Begin VB.Form frmLEVELSELECT 
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   224
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "3"
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdSHOP 
      Caption         =   "Visit the shop"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "2"
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdLEVEL 
      Caption         =   "1"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Choose a level:"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmLEVELSELECT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const NUMBEROFLEVELS = 2

Private Sub cmdLEVEL_Click(Index As Integer)
    'default of 0
    intMONSTERSONLEVEL(0) = 0
    intMONSTERSONLEVEL(1) = 0
    intMONSTERSONLEVEL(2) = 0
    intMONSTERSONLEVEL(3) = 0
    intMONSTERSONLEVEL(4) = 0
    intMONSTERSONLEVEL(5) = 0
    intMONSTERSONLEVEL(6) = 0
    Select Case Index + 1
        Case 1
            intMONSTERSONLEVEL(0) = 10
        Case 2
            intMONSTERSONLEVEL(0) = 20
            intMONSTERSONLEVEL(1) = 5
        Case 3
            intMONSTERSONLEVEL(0) = 1
            intMONSTERSONLEVEL(1) = 1
            intMONSTERSONLEVEL(2) = 1
            intMONSTERSONLEVEL(3) = 1
            intMONSTERSONLEVEL(4) = 1
            intMONSTERSONLEVEL(5) = 1
            intMONSTERSONLEVEL(6) = 1
    End Select
    frmATTACK.Show
    Unload frmLEVELSELECT
End Sub

Private Sub cmdSHOP_Click()
    frmSTORE.Show
    Unload frmLEVELSELECT
End Sub
